// Copyright © 2018 Paddy Xu
// 
// This file is part of QuickLook program.
// 
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using System.IO;
using System.Linq;
using System.Windows;
using QuickLook.Common.Plugin;

namespace QuickLook.Plugin.OfficeViewer
{
    public class Plugin : IViewer
    {
        private readonly string[] _formats =
            {".doc", ".docm", ".docx", ".rtf", ".xls", ".xlsx", ".xlsm", ".pptx", ".pptm", ".potx", ".potm"};

        public int Priority => 0;

        public void Init()
        {
            SyncfusionKey.Register();
        }

        public bool CanHandle(string path)
        {
            return !Directory.Exists(path) && _formats.Contains(Path.GetExtension(path).ToLower());
        }

        public void Prepare(string path, ContextObject context)
        {
            context.SetPreferredSizeFit(new Size {Width = 1920, Height = 1440}, 0.9);
        }

        public void View(string path, ContextObject context)
        {
            var viewer = SyncfusionControl.Open(path);

            context.ViewerContent = viewer;
            context.Title = $"{Path.GetFileName(path)}";

            context.IsBusy = false;
        }

        public void Cleanup()
        {
        }
    }
}