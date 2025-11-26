using System;
using System.IO;

namespace analyzer.Core
{
    internal sealed class RelaxAnalyzerConfig
    {
        private RelaxAnalyzerConfig(string baseDirectory)
        {
            BaseDirectory = baseDirectory ?? AppDomain.CurrentDomain.BaseDirectory;
            ProjectDirectory = BaseDirectory;
            TypeCsvPath = Path.Combine(BaseDirectory, "rule", "type.csv");
        }

        public string BaseDirectory { get; }

        public string ProjectDirectory { get; private set; }

        public string TypeCsvPath { get; private set; }

        public static RelaxAnalyzerConfig Load()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var config = new RelaxAnalyzerConfig(baseDir);
            var iniPath = Path.Combine(baseDir, "config.ini");
            if (!File.Exists(iniPath))
            {
                return config;
            }

            foreach (var rawLine in File.ReadAllLines(iniPath))
            {
                var line = rawLine.Trim();
                if (string.IsNullOrEmpty(line) || line.StartsWith(";", StringComparison.Ordinal) || line.StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                var splitterIndex = line.IndexOf('=');
                if (splitterIndex < 0)
                {
                    continue;
                }

                var key = line.Substring(0, splitterIndex).Trim();
                var value = line.Substring(splitterIndex + 1).Trim();

                if (key.Equals("Project", StringComparison.OrdinalIgnoreCase))
                {
                    config.ProjectDirectory = ResolvePath(value, config.BaseDirectory);
                }
                else if (key.Equals("TypeCSV", StringComparison.OrdinalIgnoreCase))
                {
                    config.TypeCsvPath = ResolvePath(value, config.ProjectDirectory ?? config.BaseDirectory);
                }
            }

            return config;
        }

        private static string ResolvePath(string candidate, string fallbackRoot)
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return fallbackRoot;
            }

            if (Path.IsPathRooted(candidate))
            {
                return Path.GetFullPath(candidate);
            }

            var root = string.IsNullOrWhiteSpace(fallbackRoot) ? AppDomain.CurrentDomain.BaseDirectory : fallbackRoot;
            return Path.GetFullPath(Path.Combine(root, candidate));
        }
    }
}

