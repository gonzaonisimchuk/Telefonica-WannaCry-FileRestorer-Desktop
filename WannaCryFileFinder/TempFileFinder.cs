namespace WannaCryFileFinder
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;

    public static class TempFileFinder
    {
        private const string TEMP_FILE_PATTERN = "*.WNCRYT";
        private const string TEMP_DIRECTORY_NON_SYSTEM_DISK = @"$RECYCLE\";
        public static IEnumerable<string> Find()
        {
            var foldersToSearch = new List<string>
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),@"temp\")
            };
            foldersToSearch.AddRange(DriveInfo.GetDrives().Select(p => Path.Combine(p.RootDirectory.FullName, TEMP_DIRECTORY_NON_SYSTEM_DISK)));

            var queue = new ConcurrentQueue<string>();

            var findTasks = (from folder in foldersToSearch
                             where Directory.Exists(folder)
                             select Task.Factory.StartNew(() => Find(folder, true).All(p =>
                             {
                                 queue.Enqueue(p);
                                 return true;
                             }))).Cast<Task>().ToList();

            while (queue.Count > 0 || findTasks.Any(p => !p.IsCompleted))
            {
                string file;
                if (queue.TryDequeue(out file))
                    yield return file;
                else
                    Task.Delay(500).Wait();
            }
        }

        public static void CopyRecognizedFilesTo(string destinationPath, bool overwriteFiles, params string[] files)
        {
            Parallel.ForEach(files, new ParallelOptions() { MaxDegreeOfParallelism = Environment.ProcessorCount }, (file) =>
                  {
                      try
                      {
                          File.Copy(file, Path.Combine(destinationPath, Path.ChangeExtension(Path.GetFileName(file), MimeHelper.GetExtensionByContent(file))), overwriteFiles);
                      }
                      catch (Exception)
                      {
                          // ignored
                      }
                  });
        }

        private static IEnumerable<string> Find(string path, bool recursive)
        {
            return GetDirectoryFiles(path, TEMP_FILE_PATTERN, recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
        }

        private static IEnumerable<string> GetDirectoryFiles(string rootPath, string patternMatch, SearchOption searchOption)
        {
            var foundFiles = Enumerable.Empty<string>();

            try
            {
                foundFiles = foundFiles.Concat(Directory.EnumerateFiles(rootPath, patternMatch));
            }
            catch (UnauthorizedAccessException) { }

            if (searchOption != SearchOption.AllDirectories)
                return foundFiles;
            try
            {
                var subDirs = Directory.EnumerateDirectories(rootPath);
                foreach (var dir in subDirs)
                {
                    foundFiles = foundFiles.Concat(GetDirectoryFiles(dir, patternMatch, searchOption));
                }
            }
            catch (UnauthorizedAccessException) { }
            catch (PathTooLongException) { }

            return foundFiles;
        }
    }
}
