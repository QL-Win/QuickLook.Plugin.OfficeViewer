using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace QuickLook.Plugin.OfficeViewer
{
    internal static class LegacyPowerPointConverter
    {
        public static string CacheRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "QuickLook", "OfficeViewer", "LegacyPowerPoint", "v1");

        public static string Convert(string source, CancellationToken cancellation)
        {
            source = Path.GetFullPath(source);
            if (!string.Equals(Path.GetExtension(source), ".ppt", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Only .ppt files are supported.");
            cancellation.ThrowIfCancellationRequested();
            // Keep the source locked against writes while hashing and copying.
            using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var sha = SHA256.Create())
            {
                string key = BitConverter.ToString(sha.ComputeHash(input)).Replace("-", "").ToLowerInvariant();
                string folder = Path.Combine(CacheRoot, key);
                string result = Path.Combine(folder, "preview.pdf");
                if (ValidPdf(result)) return result;
                // Separate QuickLook instances may request the same document.
                using (var gate = new Mutex(false, "Local\\QuickLook.OfficeViewer.LegacyPowerPoint." + key))
                {
                    bool acquired = false;
                    try
                    {
                        while (!acquired)
                        {
                            cancellation.ThrowIfCancellationRequested();
                            try { acquired = gate.WaitOne(200); }
                            catch (AbandonedMutexException) { acquired = true; }
                        }
                        if (ValidPdf(result)) return result;
                        Directory.CreateDirectory(folder);
                        // LibreOffice profile files can exceed Windows' path limit if
                        // nested under a long user path and a 64-character cache key.
                        string job = Path.Combine(Path.GetTempPath(), "QuickLook-PptJobs", Guid.NewGuid().ToString("N"));
                        Directory.CreateDirectory(job);
                        try
                        {
                            string snapshot = Path.Combine(job, "source.ppt");
                            input.Position = 0;
                            using (var copy = File.Create(snapshot)) input.CopyTo(copy);
                            cancellation.ThrowIfCancellationRequested();
                            RunLibreOffice(snapshot, job);
                            string generated = Path.Combine(job, "source.pdf");
                            if (!ValidPdf(generated))
                                throw new IOException("LibreOffice did not produce a readable PDF preview.");
                            // Publish only completed files. Never overwrite a valid cached PDF.
                            string pending = Path.Combine(folder, "pending-" + Guid.NewGuid().ToString("N") + ".pdf");
                            try
                            {
                                File.Copy(generated, pending);
                                if (File.Exists(result)) File.Delete(result);
                                File.Move(pending, result);
                            }
                            finally { if (File.Exists(pending)) File.Delete(pending); }
                            return result;
                        }
                        finally
                        {
                            // job is a private GUID directory created by this operation.
                            try { Directory.Delete(job, true); } catch (IOException) { }
                            catch (UnauthorizedAccessException) { }
                        }
                    }
                    finally { if (acquired) gate.ReleaseMutex(); }
                }
            }
        }

        internal static bool IsAvailable { get { return LibreOffice() != null; } }

        private static string LibreOffice()
        {
            foreach (string root in new[] {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) })
            {
                string exe = Path.Combine(root, "LibreOffice", "program", "soffice.com");
                if (File.Exists(exe)) return exe;
            }
            return null;
        }

        private static void RunLibreOffice(string source, string job)
        {
            string profile = new Uri(Path.Combine(job, "profile") + Path.DirectorySeparatorChar).AbsoluteUri;
            var info = new ProcessStartInfo
            {
                FileName = LibreOffice() ?? throw new FileNotFoundException("Install LibreOffice to preview legacy .ppt files."),
                Arguments = Quote("-env:UserInstallation=" + profile) +
                    " --headless --nologo --nodefault --norestore --convert-to " +
                    Quote("pdf:impress_pdf_Export") + " --outdir " + Quote(job) + " " + Quote(source),
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using (var process = Process.Start(info))
            {
                Task<string> output = process.StandardOutput.ReadToEndAsync();
                Task<string> errors = process.StandardError.ReadToEndAsync();
                var timer = Stopwatch.StartNew();
                // Let a started conversion finish on preview close so its worker child
                // can exit normally and the next opening can reuse the completed cache.
                // A hung process is terminated as a tree, never other LibreOffice sessions.
                while (!process.WaitForExit(200))
                {
                    if (timer.Elapsed > TimeSpan.FromMinutes(3))
                    {
                        KillTree(process);
                        throw new TimeoutException("Legacy PowerPoint conversion timed out.");
                    }
                }
                process.WaitForExit();
                if (process.ExitCode != 0)
                    throw new IOException("LibreOffice conversion failed (" + process.ExitCode + "): " +
                        errors.GetAwaiter().GetResult() + " " + output.GetAwaiter().GetResult());
                output.GetAwaiter().GetResult();
                errors.GetAwaiter().GetResult();
            }
        }

        private static void KillTree(Process process)
        {
            if (process.HasExited) return;
            string taskkill = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "taskkill.exe");
            using (var killer = Process.Start(new ProcessStartInfo {
                FileName = taskkill, Arguments = "/PID " + process.Id + " /T /F",
                UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden }))
                killer.WaitForExit(10000);
        }

        private static string Quote(string value)
        {
            // Windows paths cannot contain quotes; no shell is involved.
            if (value.IndexOf('"') >= 0) throw new ArgumentException("Invalid file path.");
            return "\"" + value + "\"";
        }

        private static bool ValidPdf(string path)
        {
            if (!File.Exists(path)) return false;
            using (var stream = File.OpenRead(path))
            {
                byte[] header = new byte[5];
                return stream.Read(header, 0, 5) == 5 && Encoding.ASCII.GetString(header) == "%PDF-";
            }
        }
    }
}
