using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

using Siemens.Engineering;
using Siemens.Engineering.AddIn;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.AddIn.Workflow;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HmiUnified.HmiAlarm;
using Siemens.Engineering.HmiUnified.HmiTags;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;

using SimaticSW = Siemens.Engineering.SW.Blocks;

using System.Text.RegularExpressions;
using System.Globalization;
using TIA_Extract.Utility;
using Extract.Core;
using System.Diagnostics;
using Siemens.Engineering.HmiUnified.HmiConnections;
using Extract.Models;
using Siemens.Engineering.HmiUnified.UI.Widgets;
using Siemens.Engineering.SW.Tags;

namespace TIA_Extract
{
    public class AlarmsTextsAddIn : ContextMenuAddIn
    {
        /// <summary>
        /// The instance of TIA Portal in which the Add-In works.
        /// <para>Enables Add-In to interact with TIA Portal.</para>
        /// </summary>
        private readonly TiaPortal _tiaPortal;

        private Settings _settings;

        /// <summary>
        /// The display name of the Add-In.
        /// </summary>
        private const string s_DisplayNameOfAddIn = "Alarms HMI";

        private const string userFolder = "UserFiles";

        private FeedbackContext _feedbackContext;

        /// <summary>
        /// Current application GUI - window
        /// </summary>
        public static Extract_UI.MainWindow MainWindow;

        /// <summary>
        /// The constructor of the AddIn.
        /// Creates an instance of the class AddIn.
        /// <para>- Called from AddInProvider when the first right-click is performed in TIA Portal.</para>
        /// <para>- The base class' constructor of AddIn will also be executed.</para>
        /// </summary>
        /// <param name="tiaPortal">
        /// Represents the instance of TIA Portal in which the Add-In will work.
        /// </param>
        public AlarmsTextsAddIn(TiaPortal tiaPortal) : base(s_DisplayNameOfAddIn)
        {
            _tiaPortal = tiaPortal;

            //#if DEBUG
            //            Core.Properties.Resources.Culture = new CultureInfo("en-US");
            //#else
            //            TiaPortalSettingsFolder generalSettingsFolder = _tiaPortal.SettingsFolders.Find("General");
            //            TiaPortalSetting UILanguageSetting = generalSettingsFolder.Settings.Find("UserInterfaceLanguage");
            //            Core.Properties.Resources.Culture = UILanguageSetting.Value as CultureInfo;
            //#endif
        }

        /// <summary>
        /// The method is provided to create a submenu of the Add-In's context menu item.
        /// Called when a mouse-over is performed on the Add-In's context menu item.
        /// </summary>
        /// <param name="addInRootSubmenu">
        /// Submenu of the Add-In's context menu item.
        /// </param>
        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>(Extract.Core.Properties.Resources.ContextMenu_BuildAlarms, OnGenerateAlarms, OnCanGenerate);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>(Extract.Core.Properties.Resources.ContextMenu_BuildFolderTags, OnGenerateTagsFolders, OnCanGenerate);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>(Extract.Core.Properties.Resources.ContextMenu_BuildTags, OnGenerateHmiTags, OnCanGenerate);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>(Extract.Core.Properties.Resources.ContextMenu_Settings, OnSettings);
        }

        /// <summary>
        /// Called when mouse is over the context menu item 'Action 1'.
        /// The returned value will be used to enable or disable it.
        /// </summary>
        /// <param name="menuSelectionProvider">
        /// Here, the same generic type as was used in addInRootSubmenu.Items.AddActionItem must be used
        /// (here it has to be IEngineeringObject)
        /// </param>
        private MenuStatus OnCanGenerate(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            // MenuStatus
            //  Enabled  = Visible
            //  Disabled = Visible but not executable
            //  Hidden   = Item will not be shown

            if (menuSelectionProvider.GetSelection<SimaticSW.GlobalDB>().Any() ||
                menuSelectionProvider.GetSelection<SimaticSW.PlcBlockUserGroup>().Any() //||
                //menuSelectionProvider.GetSelection<TagTable>().Any()
               )
                return MenuStatus.Enabled;
            else
                return MenuStatus.Hidden;
        }

        private void OnGenerateAlarms(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            GetFeedbackContext();
            _settings = Settings.Load();

            if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary) is null && _tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project is null)
            {
                _feedbackContext.Log(NotificationIcon.Information, Extract.Core.Properties.Resources.Feedback_NoProject_Text);
                return;
            }

            // Gets the TiaPortal version from the project filename extension
            //Project project = _tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary);
            var project = (ITransactionSupport)_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary) ?? (_tiaPortal.LocalSessions?.FirstOrDefault(x => x.Project.IsPrimary)?.Project);

            var projectPath = _tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.Path ?? _tiaPortal.LocalSessions?.FirstOrDefault(x => x.Project.IsPrimary)?.Project?.Path;
            //var _tiaVersion = Regex.Match(projectPath.Extension.Replace("_", "."), @"\d.+", RegexOptions.IgnoreCase).Value;

            using (var exclusiveAccess = _tiaPortal.ExclusiveAccess())
            {
                var projectConnexions = GetConnexionBetweenPlcHmi(menuSelectionProvider, exclusiveAccess);

                var referenceCulture = new List<CultureInfo>();

                if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary) != null)
                    referenceCulture.Add(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.LanguageSettings?.ReferenceLanguage?.Culture);
                if (_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary) != null)
                    referenceCulture.Add(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.LanguageSettings.ReferenceLanguage.Culture);

                var activeLanguages = new List<Language>();
                if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.LanguageSettings?.ActiveLanguages?.Count > 0)
                    activeLanguages.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary).LanguageSettings.ActiveLanguages);

                if (_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project?.LanguageSettings?.ActiveLanguages?.Count > 0)
                    activeLanguages.AddRange(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.LanguageSettings.ActiveLanguages);

                var tagFolders = new List<Extract.Models.UnifiedTagFolder>();
                var tags = new List<Extract.Models.UnifiedTag>();
                var alarms = new List<Extract.Models.UnifiedAlarm>();

                foreach (var dataBlockConnexion in projectConnexions)
                {
                    //var transact = exclusiveAccess.Transaction(project, $"Alarm create {dataBlockConnexion.PlcBlock.Name}");
                    (var rstTagsFolder, var rstTags, var rstAlarms) = ExtractTagsFromDataBlock(exclusiveAccess, dataBlockConnexion.Hmi, dataBlockConnexion.PlcBlock, dataBlockConnexion.Connexion, projectPath.Directory.FullName, activeLanguages);
                    if(rstAlarms != null) alarms.AddRange(rstAlarms);
                    _feedbackContext.Log(NotificationIcon.Success, Extract.Core.Properties.Resources.Feedback_AlarmsUpdated_Text);
                    //transact.CommitOnDispose();
                }

                for (var i = 0; i < alarms.Count(); i++)
                {
                    exclusiveAccess.Text = $"{Extract.Core.Properties.Resources.BuildAlarms_BuildAlarmFrom} {alarms[i].Tagname} ({i}/{alarms.Count()})";
                    using(var transact = exclusiveAccess.Transaction(project, $"Create tag {alarms[i].Tagname}"))
                    {
                        UnifiedAlarmCreate(alarms[i], activeLanguages);
                        transact.CommitOnDispose();
                    }
                }

                _feedbackContext.Log(NotificationIcon.Success, Extract.Core.Properties.Resources.BuildAlarmsEnded);
            }
        }

        private void OnGenerateTagsFolders(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            GetFeedbackContext();
            _settings = Settings.Load();

            if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary) is null && _tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project is null)
            {
                _feedbackContext.Log(NotificationIcon.Information, Extract.Core.Properties.Resources.Feedback_NoProject_Text);
                return;
            }

            var project = (ITransactionSupport)_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary) ?? (_tiaPortal.LocalSessions?.FirstOrDefault(x => x.Project.IsPrimary)?.Project);

            var projectPath = _tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.Path ?? _tiaPortal.LocalSessions?.FirstOrDefault(x => x.Project.IsPrimary)?.Project?.Path;
            //var _tiaVersion = Regex.Match(projectPath.Extension.Replace("_", "."), @"\d.+", RegexOptions.IgnoreCase).Value;

            using (var exclusiveAccess = _tiaPortal.ExclusiveAccess())
            {
                var projectConnexions = GetConnexionBetweenPlcHmi(menuSelectionProvider, exclusiveAccess);

                var referenceCulture = new List<CultureInfo>();

                if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary) != null)
                    referenceCulture.Add(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.LanguageSettings?.ReferenceLanguage?.Culture);
                if (_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary) != null)
                    referenceCulture.Add(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.LanguageSettings.ReferenceLanguage.Culture);

                var activeLanguages = new List<Language>();
                if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.LanguageSettings?.ActiveLanguages?.Count > 0)
                    activeLanguages.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary).LanguageSettings.ActiveLanguages);

                if (_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project?.LanguageSettings?.ActiveLanguages?.Count > 0)
                    activeLanguages.AddRange(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.LanguageSettings.ActiveLanguages);

                var tagFolders = new List<Extract.Models.UnifiedTagFolder>();
                var tags = new List<Extract.Models.UnifiedTag>();
                var alarms = new List<Extract.Models.UnifiedAlarm>();

                foreach (var dataBlockConnexion in projectConnexions)
                {
                    (var rstTagsFolder, var rstTags, var rstAlarms) = ExtractTagsFromDataBlock(exclusiveAccess, dataBlockConnexion.Hmi, dataBlockConnexion.PlcBlock, dataBlockConnexion.Connexion, projectPath.Directory.FullName, activeLanguages);
                    if (rstTags != null) tags.AddRange(rstTags);
                    _feedbackContext.Log(NotificationIcon.Success, Extract.Core.Properties.Resources.Feedback_AlarmsUpdated_Text);
                }

                for (var i = 0; i < tagFolders.Count(); i++)
                {
                    exclusiveAccess.Text = $"{Extract.Core.Properties.Resources.BuildAlarms_BuildAlarmFrom} {tagFolders[i].Name} ({i}/{tags.Count()})";
                    using (var transact = exclusiveAccess.Transaction(project, $"Create tag folder {tags[i].Tagname}"))
                    {
                        UnifiedTagFolders(tagFolders[i]);
                        transact.CommitOnDispose();
                    }
                }


                _feedbackContext.Log(NotificationIcon.Success, Extract.Core.Properties.Resources.BuildAlarmsEnded);
            }
        }

        private void OnGenerateHmiTags(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            GetFeedbackContext();
            _settings = Settings.Load();

            if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary) is null && _tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project is null)
            {
                _feedbackContext.Log(NotificationIcon.Information, Extract.Core.Properties.Resources.Feedback_NoProject_Text);
                return;
            }

            var project = (ITransactionSupport)_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary) ?? (_tiaPortal.LocalSessions?.FirstOrDefault(x => x.Project.IsPrimary)?.Project);

            var projectPath = _tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.Path ?? _tiaPortal.LocalSessions?.FirstOrDefault(x => x.Project.IsPrimary)?.Project?.Path;
            //var _tiaVersion = Regex.Match(projectPath.Extension.Replace("_", "."), @"\d.+", RegexOptions.IgnoreCase).Value;

            using (var exclusiveAccess = _tiaPortal.ExclusiveAccess())
            {
                var projectConnexions = GetConnexionBetweenPlcHmi(menuSelectionProvider, exclusiveAccess);

                var referenceCulture = new List<CultureInfo>();

                if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary) != null)
                    referenceCulture.Add(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.LanguageSettings?.ReferenceLanguage?.Culture);
                if (_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary) != null)
                    referenceCulture.Add(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.LanguageSettings.ReferenceLanguage.Culture);

                var activeLanguages = new List<Language>();
                if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.LanguageSettings?.ActiveLanguages?.Count > 0)
                    activeLanguages.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary).LanguageSettings.ActiveLanguages);

                if (_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project?.LanguageSettings?.ActiveLanguages?.Count > 0)
                    activeLanguages.AddRange(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.LanguageSettings.ActiveLanguages);

                var tagFolders = new List<Extract.Models.UnifiedTagFolder>();
                var tags = new List<Extract.Models.UnifiedTag>();
                var alarms = new List<Extract.Models.UnifiedAlarm>();

                foreach (var dataBlockConnexion in projectConnexions)
                {
                    (var rstTagsFolder, var rstTags, var rstAlarms) = ExtractTagsFromDataBlock(exclusiveAccess, dataBlockConnexion.Hmi, dataBlockConnexion.PlcBlock, dataBlockConnexion.Connexion, projectPath.Directory.FullName, activeLanguages);
                    if (rstTags != null) tags.AddRange(rstTags);
                    _feedbackContext.Log(NotificationIcon.Success, Extract.Core.Properties.Resources.Feedback_AlarmsUpdated_Text);   
                }

                for (var i = 0; i < tags.Count(); i++)
                {
                    exclusiveAccess.Text = $"{Extract.Core.Properties.Resources.BuildAlarms_BuildAlarmFrom} {tags[i].Tagname} ({i}/{tags.Count()})";
                    using (var transact = exclusiveAccess.Transaction(project, $"Create tag {tags[i].Tagname}"))
                    {
                        UnifiedTag(tags[i]);
                        transact.CommitOnDispose();
                    }
                }


                _feedbackContext.Log(NotificationIcon.Success, Extract.Core.Properties.Resources.BuildAlarmsEnded);
            }
        }

        private void BuildTagFolder(ExclusiveAccess exclusiveAccess)
        {

        }

        private List<Extract.Models.ProjectConnexion> GetConnexionBetweenPlcHmi(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider, ExclusiveAccess exclusiveAccess)
        {
            var projectConnexions = new List<Extract.Models.ProjectConnexion>();
            var devices = GetAllDevices();

            foreach (var plcBlock in menuSelectionProvider.GetSelection<IEngineeringObject>())
            {
                if (exclusiveAccess.IsCancellationRequested)
                {
                    _feedbackContext.Log(NotificationIcon.Information, Extract.Core.Properties.Resources.CancelByUser);
                    return null;
                }

                switch (plcBlock)
                {
                    case SimaticSW.GlobalDB globalDb:
                        if (globalDb.Name.EndsWith(_settings.BlockExtension))
                        {
                            var plcSoft = GetDeviceSoftware<PlcSoftware>(globalDb);
                            foreach (var device in devices)
                            {
                                foreach (var deviceItem in device.DeviceItems)
                                {
                                    var softContainer = deviceItem.GetService<SoftwareContainer>();
                                    switch (softContainer?.Software)
                                    {
                                        case HmiTarget hmi:
                                            break;
                                        case HmiSoftware hmiUnified:
                                            //if (GetDeviceSoftware<HmiSoftware>(device) is HmiSoftware hmiUnified && hmiUnified.Connections.Any(a => a.Partner == plcSoft.Name))
                                            if (hmiUnified.Connections.Any(a => a.Partner == plcSoft.Name))
                                            {
                                                foreach (var connection in hmiUnified.Connections.Where(a => a.Partner == plcSoft.Name))
                                                {
                                                    projectConnexions.Add(new Extract.Models.ProjectConnexion(hmiUnified, globalDb, connection));
                                                    //dataBlocks.Add(globalDb);
                                                    Debug.WriteLine($"Add data block {globalDb.Name}");
                                                }
                                            }
                                            break;
                                    }

                                }
                            }

                        }
                        break;
                    case SimaticSW.PlcBlockUserGroup blockGroup:
                        var plcSoft2 = GetDeviceSoftware<PlcSoftware>(blockGroup);

                        var tempGroupList = new List<SimaticSW.PlcBlockUserGroup> { blockGroup };
                        do
                        {
                            if (tempGroupList[0].Blocks.Where(bloc => bloc is SimaticSW.GlobalDB).Cast<SimaticSW.GlobalDB>() is IEnumerable<SimaticSW.GlobalDB> globalDBs)
                            {
                                if (globalDBs.Any(block => block.Name.EndsWith(_settings.BlockExtension)))
                                {
                                    foreach (var plcDataBlock in globalDBs.Where(block => block.Name.EndsWith(_settings.BlockExtension)))
                                    {
                                        //dataBlocks.Add(plcDataBlock);
                                        foreach (var device in devices)
                                        {
                                            foreach (var deviceItem in device.DeviceItems)
                                            {
                                                var softContainer = deviceItem.GetService<SoftwareContainer>();
                                                switch (softContainer?.Software)
                                                {
                                                    case HmiTarget hmi:
                                                        break;
                                                    case HmiSoftware hmiUnified:
                                                        if (hmiUnified.Connections.Any(a => a.Partner == plcSoft2.Name))
                                                        {
                                                            foreach (var connection in hmiUnified.Connections.Where(a => a.Partner == plcSoft2.Name))
                                                            {
                                                                projectConnexions.Add(new Extract.Models.ProjectConnexion(hmiUnified, plcDataBlock, connection));
                                                                //dataBlocks.Add(globalDb);
                                                                Debug.WriteLine($"Add data block {plcDataBlock.Name}");
                                                            }
                                                        }
                                                        break;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    _feedbackContext.Log(NotificationIcon.Information, $"{Extract.Core.Properties.Resources.EmptyGlobalDb} {blockGroup.Name}");
                                }
                            }

                            if (tempGroupList[0].Groups.Count > 0)
                                tempGroupList.AddRange(tempGroupList[0].Groups);

                            tempGroupList.RemoveAt(0);

                        } while (tempGroupList.Count > 0);
                        break;
                    case TagTable hmiTagTable:
                        //var hmiSoft = GetDeviceSoftware<HmiSoftware>(hmiTagTable);
                        //foreach (var connection in hmiSoft.Connections)
                        //{
                        //    var devices = new List<Device>();
                        //    devices.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary).Devices);
                        //    devices.AddRange(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.Devices);

                        //    var deviceGroup = new List<DeviceUserGroup>();
                        //    deviceGroup.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary).DeviceGroups);
                        //    deviceGroup.AddRange(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.DeviceGroups);
                        //    do
                        //    {
                        //        if (deviceGroup.Count > 0)
                        //        {
                        //            devices.AddRange(deviceGroup[0].Devices);
                        //            deviceGroup.AddRange(deviceGroup[0].Groups);
                        //        }
                        //        deviceGroup.RemoveAt(0);
                        //    } while (deviceGroup.Count > 0);

                        //    foreach (var device in devices)
                        //    {
                        //        var plcSoftware = GetDeviceSoftware<PlcSoftware>(device);
                        //        if (plcSoftware.Name == connection.Partner)
                        //        {
                        //            var group = plcSoftware.BlockGroup.Groups.Find(hmiTagTable.Name);
                        //            foreach (var plcGroupBlock in group.Blocks.Where(bloc => bloc.Name.EndsWith(_settings.BlockExtension)))
                        //                if (plcGroupBlock is SimaticSW.GlobalDB dB)
                        //                {
                        //                    BuildAlarms(exclusiveAccess, dB, projectPath.Directory.FullName);
                        //                    _feedbackContext.Log(NotificationIcon.Success, Extract.Core.Properties.Resources.Feedback_AlarmsUpdated_Text);
                        //                }
                        //        }
                        //    }
                        //}
                        break;
                }
            }
            return projectConnexions;
        }

        private (List<Extract.Models.UnifiedTagFolder>, List<Extract.Models.UnifiedTag>, List<Extract.Models.UnifiedAlarm>) ExtractTagsFromDataBlock(ExclusiveAccess exclusiveAccess, HmiSoftware hmiSoft, SimaticSW.PlcBlock globalDB, HmiConnection hmiConnection, 
            string projectDirectoryPath, List<Language> activeLanguages)
        {
            _feedbackContext.Log(NotificationIcon.Information, $"{Extract.Core.Properties.Resources.BuildAlarms_ExtractAlarmForm} {globalDB.Name}");
            exclusiveAccess.Text = $"{Extract.Core.Properties.Resources.BuildAlarms_BuildAlarmFrom} {globalDB.Name}";

            var tagFolders = new List<Extract.Models.UnifiedTagFolder>();
            var tags = new List<Extract.Models.UnifiedTag>();
            var alarms = new List<Extract.Models.UnifiedAlarm>();

           

            if (globalDB.Name.EndsWith(_settings.BlockExtension))
            {
                var folderName = string.Empty;
                if (globalDB.Parent is SimaticSW.PlcBlockUserGroup group)
                    folderName = group.Name;

                if (!globalDB.IsConsistent)
                {
                    var comp = globalDB.GetService<ICompilable>();
                    if (comp.Compile().State != CompilerResultState.Success)
                    {
                        _feedbackContext.Log(NotificationIcon.Error, $"{globalDB.Name} {Extract.Core.Properties.Resources.BuildAlarms_CannotCompile}");
                        return (null, null, null);
                    }
                }

                var path = Path.Combine(projectDirectoryPath, userFolder, globalDB.Name + ".xml");
                Util.ExportBlock(globalDB, path, _feedbackContext);

                Debug.WriteLine($"Data block {globalDB.Name} is exported");

                var serializer = new XmlSerializer(typeof(SimaticML.Document));
                var myFile = new FileStream(path, FileMode.Open);
                if (serializer.Deserialize(myFile) is SimaticML.Document document)
                    foreach (var Db in document.Items.Where(sw => sw is SimaticML.SW.Blocks.GlobalDB).Cast<SimaticML.SW.Blocks.GlobalDB>())
                    {
                        if (exclusiveAccess.IsCancellationRequested)
                        {
                            _feedbackContext.Log(NotificationIcon.Information, Extract.Core.Properties.Resources.TaskGenerateCancel);
                            return (null, null, null);
                        }

                        var exportGlobalDataBlock = Db.AttributeList.Interface.Sections.FirstOrDefault(section => section.Name == SimaticML.SW.Common.SectionName_TE.Static);
                        if (exportGlobalDataBlock is null) break;

                        var alarmsClassName = _settings.DefaultAlarmsClass;
                        foreach (var exportedMember in exportGlobalDataBlock.Member)
                        {
                            foreach (var member in GetInterfaceMember(exportedMember))
                            {
                                switch (member.Type)
                                {
                                    case "Struct":
                                        foreach (var lang in activeLanguages)
                                        {
                                            if (member.Description.ContainsKey(lang.Culture.Name))
                                            {
                                                var txt = member.Description[lang.Culture.Name];
                                                if (!string.IsNullOrEmpty(txt))
                                                {
                                                    var match = Regex.Match(txt, @"\[AlarmClass=(.+?)\]", RegexOptions.IgnoreCase);
                                                    alarmsClassName = GetAlarmsClassName(hmiSoft, match.Groups[1].Value, member.Name, _settings.DefaultAlarmsClass);

                                                    Debug.WriteLine($"Class alarm {alarmsClassName} from {globalDB.Name}");
                                                    break;
                                                }
                                            }
                                        }

                                        break;
                                    case "Bool":
                                        tags.Add(new Extract.Models.UnifiedTag
                                        {
                                            Hmi = hmiSoft,
                                            Connexion = hmiConnection,
                                            PlcTag = $"{(globalDB.Name.Contains(' ') ? $"\"{globalDB.Name}\"" : globalDB.Name)}.{member.Name}",
                                            Tagname = $"{(_settings.SimplifyTagname ? globalDB.Name.Replace(_settings.BlockExtension, string.Empty) : globalDB.Name)}_{member.Name.Replace(".", "_")}",
                                            Folder = folderName
                                        });
                                        alarms.Add(new Extract.Models.UnifiedAlarm
                                        {
                                            Hmi = hmiSoft,
                                            ClassName = alarmsClassName,
                                            Tagname = $"{(_settings.SimplifyTagname ? globalDB.Name.Replace(_settings.BlockExtension, string.Empty) : globalDB.Name)}_{member.Name.Replace(".", "_")}",
                                            Origin = globalDB.Name.Replace(_settings.BlockExtension, string.Empty),
                                            Descriptions = member.Description,
                                        });

                                        Debug.WriteLine($"Add tag {member.Name}");

                                        break;
                                }
                            }
                        }
                    }

                myFile.Close();
                File.Delete(path);
            }

            Debug.WriteLine($"{alarms.Count()} alarms founded");

            return (tagFolders, tags, alarms);

            
        }

        private void OnSettings(MenuSelectionProvider menuSelectionProvider)
        {
            _settings = Settings.Load();

            MainWindow = new Extract_UI.MainWindow
            {
                Topmost = true,
                ShowActivated = true,
                DataContext = _settings
            };

            if (MainWindow.ShowDialog() == true)
                _settings.Save();

            //if (MainWindow.ClosedByUser)
            //{
            //    System.Environment.Exit(0); // exit the addin if closed by button
            //}
        }

        private T GetDeviceSoftware<T>(IEngineeringObject engineeringObject)
        {
            var plc = engineeringObject.Parent;
            while (plc != null && plc.GetType() != typeof(PlcSoftware))
            {
                plc = plc.Parent;
            }
            return plc is null ? default : (T)Convert.ChangeType(plc, typeof(T));
        }

        private void GetFeedbackContext()
        {
            if (_feedbackContext is null)
            {
                try
                {
                    _feedbackContext = _tiaPortal.GetFeedbackContext();
                }
                catch { }
            }
        }

        private List<Device> GetAllDevices()
        {
            var devices = new List<Device>();
            var deviceGroup = new List<DeviceUserGroup>();

            if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.Devices?.Count > 0)
                devices.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.Devices);

            if (_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.Devices?.Count > 0)
                devices.AddRange(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.Devices);

            if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.DeviceGroups?.Count > 0)
                deviceGroup.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.DeviceGroups);

            if (_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.DeviceGroups?.Count > 0)
                deviceGroup.AddRange(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.DeviceGroups);

            do
            {
                if (deviceGroup.Count > 0)
                {
                    devices.AddRange(deviceGroup[0].Devices);
                    deviceGroup.AddRange(deviceGroup[0].Groups);
                    deviceGroup.RemoveAt(0);
                }
            } while (deviceGroup.Count > 0);

            return devices;
        }

        private string GetAlarmsClassName(HmiSoftware hmiUnified, string classname, string tagname, string defaultClassname)
        {
            if (hmiUnified.AlarmClasses.Find(classname) is HmiAlarmClass groupAlarmClass)
            {
                return groupAlarmClass.Name;
            }
            else if (hmiUnified.AlarmClasses.Find(tagname) is HmiAlarmClass groupAlarmClass2)
            {
                return groupAlarmClass2.Name;
            }
            else
                return hmiUnified.AlarmClasses.Find(defaultClassname).Name;
        }

        private List<InterfaceMember> GetInterfaceMember(SimaticML.SW.InterfaceSections.Member_T member, string parent = "")
        {
            var tags = new List<InterfaceMember>();
            var subTags = new List<InterfaceMember>();
            var tag = new InterfaceMember
            {
                Name = string.IsNullOrEmpty(parent) ? member.Name : $"{parent}.{member.Name}",
                Direction = DirectionMember.Static,
                Type = member.Datatype
            };

            foreach (var item in member.Items)
            {
                switch (item)
                {
                    case SimaticML.SW.Common.Comment comment:
                        foreach (var multiLangugageText in comment.MultiLanguageText)
                        {
                            if (tag.Description is null) tag.Description = new Dictionary<string, string>();
                            tag.Description.Add(multiLangugageText.Lang, multiLangugageText.Value);
                        }
                        break;
                    case SimaticML.SW.InterfaceSections.Member_T memberItem:
                        subTags.AddRange(GetInterfaceMember(memberItem, tag.Name));
                        break;
                }
            }
            tags.Add(tag);
            tags.AddRange(subTags);

            return tags;
        }

        private string UnifiedPlcTagNormalized(string tagname)
        {
            var compose = tagname.Split('.');
            string plcTag = compose.FirstOrDefault();
            foreach (var item in compose.Skip(1))
            {
                if (item.Contains(" "))
                    plcTag += $".\"{item}\"";
                else
                    plcTag += $".{item}";
            }

            return plcTag;
        }

        private void UnifiedTagFolders(Extract.Models.UnifiedTagFolder tagFolder)
        {
            var tagTables = new List<HmiTagTable>();
            tagTables.AddRange(tagFolder.Hmi.TagTables);

            var tagTablesGroup = new List<HmiTagTableGroup>();
            tagTablesGroup.AddRange(tagFolder.Hmi.TagTableGroups);
            do
            {
                if (tagTablesGroup.Count > 0)
                {
                    tagTables.AddRange(tagTablesGroup[0].TagTables);
                    tagTablesGroup.AddRange(tagTablesGroup[0].Groups);
                    tagTablesGroup.RemoveAt(0);
                }
            } while (tagTablesGroup.Count > 0);

            if (!tagTables.Any(a => a.Name == tagFolder.Name))
                tagFolder.Hmi.TagTables.Create(tagFolder.Name);
        }

        private void UnifiedTag(UnifiedTag unifiedTag)
        {
            //using (var transact = ExclusiveAccess?.Transaction(Project, $"Build {unifiedTag.Tagname}"))
            {
                var plcTag = UnifiedPlcTagNormalized(unifiedTag.PlcTag);

                //var hmiTags = unifiedTag.Hmi.Tags.Where(c => c.PlcTag == plcTag);
                //foreach (var hmiTag in hmiTags)
                //{
                //    var tempAlarms = unifiedTag.Hmi.DiscreteAlarms.Where(w => w.RaisedStateTag == hmiTag.Name).Select(s => s.Name);
                //    foreach (var alarm in tempAlarms)
                //    {
                //        unifiedTag.Hmi.DiscreteAlarms.Single(s => s.Name == alarm)?.Delete();
                //    }

                //    if (hmiTags.Count() > 1)
                //    {
                //        hmiTag.Delete();
                //    }
                //}

                var tagDeleted = new List<string>();

                if(unifiedTag.Hmi.Tags.Count(c => c.PlcTag == plcTag) > 1)
                {
                    var iTag = 0;
                    do
                    {
                        if (unifiedTag.Hmi.Tags[iTag].PlcName.Equals(plcTag))
                        {
                            tagDeleted.Add(unifiedTag.Hmi.Tags[iTag].Name);
                            unifiedTag.Hmi.Tags[iTag].Delete();
                        }
                        else
                            iTag++;
                    } while (iTag < unifiedTag.Hmi.Tags.Count);
                }
                else
                {
                    tagDeleted.AddRange(unifiedTag.Hmi.Tags.Where(w => w.PlcTag == plcTag).Select(s => s.Name));
                }

                int iAlarm = 0;
                do
                {
                    if(tagDeleted.Contains(unifiedTag.Hmi.DiscreteAlarms[iAlarm].RaisedStateTag))
                    {
                        unifiedTag.Hmi.DiscreteAlarms[iAlarm].Delete();
                    }
                    else
                        iAlarm++;
                } while (iAlarm < unifiedTag.Hmi.DiscreteAlarms.Count);

                HmiTag tag;
                if (!unifiedTag.Hmi.Tags.Any(c => c.PlcTag == plcTag))
                {
                    if (string.IsNullOrEmpty(unifiedTag.Folder))
                        tag = unifiedTag.Hmi.Tags.Create(unifiedTag.Tagname);
                    else
                        tag = unifiedTag.Hmi.Tags.Create(unifiedTag.Tagname, unifiedTag.Folder);
                }
                else
                {
                    tag = unifiedTag.Hmi.Tags.Single(c => c.PlcTag == plcTag);
                }
                tag.Connection = unifiedTag.Connexion.Name;
                tag.PlcTag = plcTag;
                tag.Name = unifiedTag.Tagname;

                //transact?.CommitOnDispose();
            }
        }

        private void UnifiedAlarmCreate(Extract.Models.UnifiedAlarm alarm, List<Language> activeLanguages)
        {
            var alarms = alarm.Hmi.DiscreteAlarms.Find(alarm.Tagname) ?? alarm.Hmi.DiscreteAlarms.Create(alarm.Tagname);
            alarms.RaisedStateTag = alarm.Tagname;
            alarms.AlarmClass = alarm.ClassName;

            alarms.Origin = alarm.Origin;

            if(activeLanguages != null)
            {
                foreach (var lang in activeLanguages)
                {
                    if (alarms.EventText.Items.Single(s => s.Language.Culture.Name == lang.Culture.Name) is MultilingualTextItem multilingualText)
                    {
                        alarm.Descriptions.TryGetValue(lang.Culture.Name, out string txt);
                        multilingualText.Text = $"<body><p>{txt}</p></body>";
                    }
                }
            }

            Debug.WriteLine($"{alarm.Tagname} created");
        }

    }
}
