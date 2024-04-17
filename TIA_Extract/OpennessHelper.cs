using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Siemens.Engineering.Cax;
using Siemens.Engineering.Hmi.Communication;
using Siemens.Engineering.Hmi.Cycle;
using Siemens.Engineering.Hmi.Globalization;
using Siemens.Engineering.Hmi.RuntimeScripting;
using Siemens.Engineering.Hmi.Screen;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.Hmi.TextGraphicList;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.ExternalSources;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.SW.Types;
using Siemens.Engineering.SW.WatchAndForceTables;
using Siemens.Engineering.SW;

using Siemens.Engineering;

namespace TIA_Extract
{
    public partial class OpennessHelper
    {
        /// <summary>
        /// Imports a file under the given TIA object with the given import options
        /// </summary>
        /// <param name="destination">TIA object under which the file will be imported.</param>
        /// <param name="filePath">Path to the file</param>
        /// <param name="importOption">TIA import options</param>
        /// <exception cref="System.ArgumentNullException">Parameter is null;destination</exception>
        /// <exception cref="System.ArgumentException">Parameter is null or empty;filePath</exception>
        /// <exception cref="System.IO.IOException"></exception>
        /// <exception cref="System.UnauthorizedAccessException"></exception>
        /// <exception cref="System.IO.DirectoryNotFoundException"></exception>
        public static void ImportItem(IEngineeringCompositionOrObject destination, string filePath, ImportOptions importOption)
        {
            if (destination == null)
                throw new ArgumentNullException(nameof(destination), "Parameter is null");
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentException("Parameter is null or empty", nameof(filePath));

            var fileInfo = new FileInfo(filePath);
            filePath = fileInfo.FullName;


            if (destination is CycleComposition || destination is ConnectionComposition || destination is MultiLingualGraphicComposition
            || destination is GraphicListComposition || destination is TextListComposition)
            {
                var parameter = new Dictionary<Type, object>();
                parameter.Add(typeof(string), filePath);
                parameter.Add(typeof(ImportOptions), importOption);
                // Export prüfen
                (destination as IEngineeringComposition).Invoke("Import", parameter);
            }
            else if (destination is PlcBlockGroup)
            {
                if (Path.GetExtension(filePath).Equals(".xml"))
                    (destination as PlcBlockGroup).Blocks.Import(fileInfo, importOption);
                else
                {
                    var currentDestination = destination as IEngineeringObject;
                    while (!(currentDestination is PlcSoftware))
                    {
                        currentDestination = currentDestination.Parent;
                    }
                    var col = (currentDestination as PlcSoftware).ExternalSourceGroup.ExternalSources;

                    var sourceName = Path.GetRandomFileName();
                    sourceName = Path.ChangeExtension(sourceName, ".src");
                    var src = col.CreateFromFile(sourceName, filePath);
                    src.GenerateBlocksFromSource();
                    src.Delete();
                }
            }
            else if (destination is PlcTagTableGroup)
                (destination as PlcTagTableGroup).TagTables.Import(fileInfo, importOption);
            else if (destination is PlcWatchAndForceTableGroup)
                (destination as PlcWatchAndForceTableGroup).WatchTables.Import(fileInfo, importOption);
            else if (destination is PlcTypeGroup)
                (destination as PlcTypeGroup).Types.Import(fileInfo, importOption);
            else if (destination is PlcExternalSourceGroup)
            {
                var temp = (destination as PlcExternalSourceGroup).ExternalSources.Find(Path.GetFileName(filePath));
                if (temp != null)
                    temp.Delete();
                (destination as PlcExternalSourceGroup).ExternalSources.CreateFromFile(Path.GetFileName(filePath), filePath);
            }
            else if (destination is TagFolder)
                (destination as TagFolder).TagTables.Import(fileInfo, importOption);
            else if (destination is ScreenFolder)
                (destination as ScreenFolder).Screens.Import(fileInfo, importOption);
            else if (destination is ScreenTemplateFolder)
                (destination as ScreenTemplateFolder).ScreenTemplates.Import(fileInfo, importOption);
            else if (destination is ScreenPopupFolder)
                (destination as ScreenPopupFolder).ScreenPopups.Import(fileInfo, importOption);
            else if (destination is ScreenSlideinSystemFolder)
                (destination as ScreenSlideinSystemFolder).ScreenSlideins.Import(fileInfo, importOption);
            else if (destination is VBScriptFolder)
                (destination as VBScriptFolder).VBScripts.Import(fileInfo, importOption);
            else if (destination is ScreenGlobalElements)
                (destination.Parent as HmiTarget)?.ImportScreenGlobalElements(fileInfo, importOption);
            else if (destination is ScreenOverview)
                (destination.Parent as HmiTarget)?.ImportScreenOverview(fileInfo, importOption);
            else
            {
                Console.WriteLine("Not work");
            }

        }

        /// <summary>Imports a folder structure into the given destination</summary>
        /// <param name="targetLocation">TIA object under which the structure will be imported</param>
        /// <param name="folderPath">Path to the folder to import</param>
        /// <param name="importOption">TIA import options</param>
        /// <exception cref="System.ArgumentNullException">Parameter is null;targetLocation</exception>
        /// <exception cref="System.ArgumentException">Parameter is null or empty;folderPath</exception>
        /// <exception cref="System.IO.IOException"></exception>
        /// <exception cref="System.UnauthorizedAccessException"></exception>
        /// <exception cref="System.IO.DirectoryNotFoundException"></exception>
        //private static void ImportStructure(IEngineeringObject targetLocation, string folderPath, ImportOptions importOption)
        //{
        //    if (targetLocation == null)
        //        throw new ArgumentNullException(nameof(targetLocation), "Parameter is null");
        //    if (string.IsNullOrEmpty(folderPath))
        //        throw new ArgumentException("Parameter is null or empty", nameof(folderPath));

        //    var files = Directory.GetFiles(folderPath);
        //    var dirs = Directory.GetDirectories(folderPath);

        //    foreach (var file in files)
        //    {
        //        ImportItem(targetLocation, file, importOption);
        //    }

        //    foreach (var dir in dirs)
        //    {
        //        var subDir = GetFolderByName(targetLocation, Path.GetFileName(dir));
        //        ImportStructure(subDir as IEngineeringObject, dir, importOption);
        //    }
        //}

        //public static bool CaxImport(Project project, string filePath, ImportCaxOptions importOptions)
        //{
        //    if (project == null)
        //        throw new ArgumentNullException(nameof(project), "Parameter is null");
        //    if (String.IsNullOrEmpty(filePath))
        //        throw new ArgumentException("Parameter is null or empty", nameof(filePath));

        //    var provider = project.GetService<CaxProvider>();
        //    if (provider == null) return false;
        //    var logPath = new FileInfo(Path.Combine(Path.GetDirectoryName(filePath), "Import.log"));

        //    CaxImportOptions tmpOptions;

        //    if (Enum.TryParse(importOptions.ToString(), out tmpOptions))
        //        return provider.Import(new FileInfo(filePath), logPath, tmpOptions);

        //    return provider.Import(new FileInfo(filePath), logPath, CaxImportOptions.MoveToParkingLot);

        //}
    }
}
