

namespace AddStyles
{
    /// <summary>
    /// Класс отвечающий за работу с файловой системой (Поиск нужных файлов для работы)
    /// </summary>
    public class FileFinder
    {
        private string src_folder { get; init; }
        public string dst_folder { get; init; }
        private string current_directory { get; init; }
        public FileFinder()
        {
            //Можно улучшить добавив проверку Атаян
            current_directory = AppContext.BaseDirectory;
            Console.WriteLine("Введите номер папки-источника. ");
            src_folder = Path.Combine(current_directory, GetFolder_target());
            Console.WriteLine("Введите номер папки-назначения. ");
            dst_folder = Path.Combine(current_directory, GetFolder_target());
        }
        /// <summary>
        /// Проводит поиск директорий, соседствующих с исполняемым файлом
        /// </summary>
        /// <returns>Список названий соседствующих папок</returns>
        private List<string> Enumerate_folders()
        {
            List<string> inner_directories = Directory.GetDirectories(current_directory).ToList();

            for (int i = 0; i < inner_directories.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {Path.GetFileName(inner_directories[i])}");
            }
            return inner_directories;
        }
        /// <summary>
        /// ПОльзователь задает номер предложенной папки 
        /// </summary>
        /// <returns>соотеветствующую номеру название папки.</returns>
        private string GetFolder_target()
        {
            var inner_directories = this.Enumerate_folders();
            int number_folder = -1;

            while (true)
            {
                Console.Write("Номер: ");
                string input = Console.ReadLine();

                if (int.TryParse(input, out number_folder) &&
                    number_folder >= 1 && number_folder <= inner_directories.Count)
                {
                    break;
                }

                Console.WriteLine("Введите корректный номер из списка.");
            }
            return inner_directories[number_folder - 1];
        }

        /// <summary>
        /// Анализирует выбранную директорию-источник на файлы эксель
        /// </summary>
        /// <returns>Возвращает массив полных путей до файлов эксель</returns>
        public string[] GetFileNames()
        {
            var list_files = Directory.GetFiles(src_folder)
                    .Where(path => !Path.GetFileName(path).StartsWith("~$"))
                    .ToArray();
            foreach (var name in list_files)
            {
                Console.WriteLine(name);
            }
            var file_paths_list = list_files.Select(x =>
            {
                return $"{Path.Combine(src_folder, x)}";
            }).ToArray();
            return list_files;
        }
    }
    public static class Program
    {
        public static void Main()
        {
            var test = new FileFinder();
            var files = test.GetFileNames();
            foreach (var file in files)
            {
                var exl = new ExcelWorker(file, test.dst_folder);
                exl.main_process();
            }
        }
    }
}
