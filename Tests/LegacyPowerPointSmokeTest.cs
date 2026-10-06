using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using QuickLook.Plugin.OfficeViewer;

internal static class LegacyPowerPointSmokeTest
{
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length < 2)
                throw new ArgumentException("Usage: SmokeTest.exe CACHE_ROOT INPUT.ppt [INPUT.ppt ...]");
            LegacyPowerPointConverter.CacheRoot = Path.GetFullPath(args[0]);
            foreach (string source in new ArraySegment<string>(args, 1, args.Length - 1))
            {
                string original = Hash(source);
                var first = Task.Run(() => LegacyPowerPointConverter.Convert(source, CancellationToken.None));
                var second = Task.Run(() => LegacyPowerPointConverter.Convert(source, CancellationToken.None));
                Task.WaitAll(first, second);
                if (first.Result != second.Result) throw new Exception("Concurrent previews did not share the cache.");
                DateTime timestamp = File.GetLastWriteTimeUtc(first.Result);
                string warm = LegacyPowerPointConverter.Convert(source, CancellationToken.None);
                if (warm != first.Result || File.GetLastWriteTimeUtc(warm) != timestamp)
                    throw new Exception("A cached preview was regenerated.");
                if (Hash(source) != original) throw new Exception("The source presentation was modified.");
                using (var canceled = new CancellationTokenSource())
                {
                    canceled.Cancel();
                    try
                    {
                        LegacyPowerPointConverter.Convert(source, canceled.Token);
                        throw new Exception("Cancellation was ignored.");
                    }
                    catch (OperationCanceledException) { }
                }
                Console.WriteLine("PASS: " + Path.GetFileName(source) + " -> " + warm);
            }
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static string Hash(string path)
    {
        using (var stream = File.OpenRead(path))
        using (var sha = SHA256.Create())
            return BitConverter.ToString(sha.ComputeHash(stream));
    }
}
