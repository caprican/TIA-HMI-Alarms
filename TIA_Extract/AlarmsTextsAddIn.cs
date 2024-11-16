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
using Siemens.Engineering.SW.Blocks.Interface;
using Siemens.Engineering.SW.Units;

using System.Text.RegularExpressions;
using SimaticML;
using System.Globalization;
using TIA_Extract.Utility;

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
            //  Core.Properties.Resources.Culture = new CultureInfo("en-US");
            //#else
            //  TiaPortalSettingsFolder generalSettingsFolder = _tiaPortal.SettingsFolders.Find("General");
            //  TiaPortalSetting UILanguageSetting = generalSettingsFolder.Settings.Find("UserInterfaceLanguage");
            //  Extract.Core.Properties.Resources.Culture = UILanguageSetting.Value as CultureInfo;
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
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>(Extract.Core.Properties.Resources.ContextMenu_GlobalDb, OnGenerate, OnCanGenerate);
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

            if(menuSelectionProvider.GetSelection<SimaticSW.GlobalDB>().Any() ||
               menuSelectionProvider.GetSelection<SimaticSW.PlcBlockUserGroup>().Any() ||
               menuSelectionProvider.GetSelection<TagTable>().Any())
                return MenuStatus.Enabled;
            else
                return MenuStatus.Hidden;
        }

        private void OnGenerate(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            GetFeedbackContext();
            _settings = Settings.Load();

            if (_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary) is null && _tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project is null)
            {
                _feedbackContext.Log(NotificationIcon.Information, Extract.Core.Properties.Resources.Feedback_NoProject_Text);
                return;
            }

            // Gets the TiaPortal version from the project filename extension
            var projectPath = _tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.Path ?? _tiaPortal.LocalSessions?.FirstOrDefault(x => x.Project.IsPrimary)?.Project?.Path;
            var _tiaVersion = Regex.Match(projectPath.Extension.Replace("_", "."), @"\d.+", RegexOptions.IgnoreCase).Value;

            //var dirPath = Path.Combine(Environment.ExpandEnvironmentVariables("%TEMP%"), AppDomain.CurrentDomain.FriendlyName);
            //var dir = Directory.CreateDirectory(dirPath);
            //var _tempFilePath = dir.FullName;

            //var path = project.Path.Directory.FullName;

            using (var exclusiveAccess = _tiaPortal.ExclusiveAccess())
            {
                foreach (var plcBlock in menuSelectionProvider.GetSelection<IEngineeringObject>())
                {
                    if(exclusiveAccess.IsCancellationRequested)
                    {
                        _feedbackContext.Log(NotificationIcon.Information, Extract.Core.Properties.Resources.CancelByUser);
                        return;
                    }

                    switch(plcBlock)
                    {
                        case SimaticSW.GlobalDB globalDb:
                            BuildAlarms(exclusiveAccess, globalDb, projectPath.Directory.FullName);
                            break;
                        case SimaticSW.PlcBlockUserGroup blockGroup:

                            List<SimaticSW.PlcBlockUserGroup> tempGroupList = new List<SimaticSW.PlcBlockUserGroup> { blockGroup };
                            var asUpdateAlarms = false;
                            do
                            {
                                if (tempGroupList[0].Blocks.Where(bloc => bloc is SimaticSW.GlobalDB).Cast<SimaticSW.GlobalDB>() is IEnumerable<SimaticSW.GlobalDB> globalDBs)
                                    foreach (var plcDataBlock in globalDBs)
                                    {
                                        asUpdateAlarms = true;
                                        BuildAlarms(exclusiveAccess, plcDataBlock, projectPath.Directory.FullName);
                                        _feedbackContext.Log(NotificationIcon.Success, Extract.Core.Properties.Resources.Feedback_AlarmsUpdated_Text);
                                    }
                                
                                if (tempGroupList[0].Groups.Count > 0)
                                    tempGroupList.AddRange(tempGroupList[0].Groups);
                                
                                tempGroupList.RemoveAt(0);

                            } while (tempGroupList.Count > 0);

                            if(!asUpdateAlarms)
                            {
                                _feedbackContext.Log(NotificationIcon.Information, $"{Extract.Core.Properties.Resources.EmptyGlobalDb} {blockGroup.Name}");
                                break;
                            }

                            break;
                        case TagTable hmiTagTable:
                            var hmiSoft = GetDeviceSoftware<HmiSoftware>(hmiTagTable);
                            foreach (var connection in hmiSoft.Connections)
                            {
                                var devices = new List<Device>();
                                devices.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary).Devices);
                                devices.AddRange(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.Devices);

                                var deviceGroup = new List<DeviceUserGroup>();
                                deviceGroup.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary).DeviceGroups);
                                deviceGroup.AddRange(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.DeviceGroups);
                                do
                                {
                                    if (deviceGroup.Count > 0)
                                    {
                                        devices.AddRange(deviceGroup[0].Devices);
                                        deviceGroup.AddRange(deviceGroup[0].Groups);
                                    }
                                    deviceGroup.RemoveAt(0);
                                } while (deviceGroup.Count > 0);

                                foreach (var device in devices)
                                {
                                    var plcSoftware = GetDeviceSoftware<PlcSoftware>(device);
                                    if (plcSoftware.Name == connection.Partner)
                                    {
                                        var group = plcSoftware.BlockGroup.Groups.Find(hmiTagTable.Name);
                                        foreach (var plcGroupBlock in group.Blocks.Where(bloc => bloc.Name.EndsWith(_settings.BlockExtension)))
                                            if (plcGroupBlock is SimaticSW.GlobalDB dB)
                                            {
                                                BuildAlarms(exclusiveAccess, dB, projectPath.Directory.FullName);
                                                _feedbackContext.Log(NotificationIcon.Success, Extract.Core.Properties.Resources.Feedback_AlarmsUpdated_Text);
                                            }
                                    }
                                }
                            }
                            break;
                    }
                }
                _feedbackContext.Log(NotificationIcon.Success, Extract.Core.Properties.Resources.BuildAlarmsEnded);
            }
        }

        public void BuildAlarms(ExclusiveAccess exclusiveAccess, SimaticSW.GlobalDB globalDB, string projectDirectoryPath)
        {
            _feedbackContext.Log(NotificationIcon.Information, $"{Extract.Core.Properties.Resources.BuildAlarms_ExtractAlarmForm} {globalDB.Name}");

            exclusiveAccess.Text = $"{Extract.Core.Properties.Resources.BuildAlarms_BuildAlarmFrom} {globalDB.Name}";

            if (globalDB.Name.EndsWith(_settings.BlockExtension))
            {
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

                var folderName = string.Empty;
                if (globalDB.Parent is SimaticSW.PlcBlockUserGroup group)
                    folderName = group.Name;
                
                var plcSoft = GetDeviceSoftware<PlcSoftware>(globalDB);

                if (!globalDB.IsConsistent)
                {
                    var comp = globalDB.GetService<ICompilable>();
                    if (comp.Compile().State != CompilerResultState.Success)
                    {
                        _feedbackContext.Log(NotificationIcon.Error, $"{globalDB.Name} {Extract.Core.Properties.Resources.BuildAlarms_CannotCompile}");
                        return;
                    }
                }

                var path = Path.Combine(projectDirectoryPath, userFolder, globalDB.Name + ".xml");
                Util.ExportBlock(globalDB, path, _feedbackContext);

                var serializer = new XmlSerializer(typeof(SimaticML.Document));
                var myFile = new FileStream(path, FileMode.Open);
                if (serializer.Deserialize(myFile) is SimaticML.Document document)
                    foreach (var Db in document.Items.Where(sw => sw is SimaticML.SW.Blocks.GlobalDB).Cast<SimaticML.SW.Blocks.GlobalDB>())
                    {
                        var devices = new List<Device>();
                        var deviceGroup = new List<DeviceUserGroup>();

                        if(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.Devices?.Count > 0)
                            devices.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.Devices);
                    
                        if(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.Devices?.Count > 0)
                            devices.AddRange(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.Devices);
                    
                        if(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.DeviceGroups?.Count > 0)
                            deviceGroup.AddRange(_tiaPortal.Projects.FirstOrDefault(x => x.IsPrimary)?.DeviceGroups);
                    
                        if(_tiaPortal.LocalSessions.FirstOrDefault(x => x.Project.IsPrimary)?.Project.DeviceGroups?.Count > 0)
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

                        foreach (var device in devices)
                        {
                            foreach (var deviceItem in device.DeviceItems)
                            {
                                if (exclusiveAccess.IsCancellationRequested)
                                {
                                    _feedbackContext.Log(NotificationIcon.Information, Extract.Core.Properties.Resources.TaskGenerateCancel);
                                    return;
                                }

                                var exportDb = Db.AttributeList.Interface.Sections.FirstOrDefault(section => section.Name == SimaticML.SW.Common.SectionName_TE.Static);

                                var softContainer = deviceItem.GetService<SoftwareContainer>();
                                switch (softContainer?.Software)
                                {
                                    case HmiTarget hmi:
                                        _feedbackContext.Log(NotificationIcon.Information, $"{Extract.Core.Properties.Resources.BuildAlarms_BuilderNoCompatible} {device.Name}.");
                                        break;
                                    case HmiSoftware hmiUnified:
                                        var connexion = hmiUnified.Connections.FirstOrDefault(con => con.Partner == plcSoft.Name);

                                        if (connexion is null)
                                        {
                                            _feedbackContext.Log(NotificationIcon.Information, $" {device.Name} et {plcSoft.Name}.");
                                            break;
                                        }

                                        var alarmClass = hmiUnified.AlarmClasses.Find(_settings.DefaultAlarmsClass);

                                        var hmiTagspath = Path.Combine(projectDirectoryPath, userFolder, hmiUnified.Name + "_Tags.xml");
                                        if (File.Exists(hmiTagspath))
                                            File.Delete(hmiTagspath);
                                        hmiUnified.Tags.Export(new DirectoryInfo(hmiTagspath));

                                        _feedbackContext.Log(NotificationIcon.Success, $"export : {hmiTagspath}");

                                        foreach (var internalMember in globalDB.Interface.Members)
                                        {
                                            if (exclusiveAccess.IsCancellationRequested)
                                            {
                                                _feedbackContext.Log(NotificationIcon.Information, Extract.Core.Properties.Resources.CancelByUser);
                                                return;
                                            }

                                            if (GetExportMember(internalMember, exportDb.Member) is SimaticML.SW.InterfaceSections.Member_T exportTag)
                                            {
                                                switch (exportTag.Datatype)
                                                {
                                                    case "Struct":

                                                        foreach (var lang in referenceCulture)
                                                        {
                                                            var txt = GetCommentText(lang, exportTag);
                                                            if (!string.IsNullOrEmpty(txt))
                                                            {
                                                                var match = Regex.Match(txt, @"\[AlarmClass=(.+?)\]", RegexOptions.IgnoreCase);
                                                                if (match.Success && hmiUnified.AlarmClasses.Find(match.Groups[1].Value) is HmiAlarmClass groupAlarmClass)
                                                                {
                                                                    alarmClass = groupAlarmClass;
                                                                    break;
                                                                }
                                                                else if (hmiUnified.AlarmClasses.Find(exportTag.Name) is HmiAlarmClass groupAlarmClass2)
                                                                {
                                                                    alarmClass = groupAlarmClass2;
                                                                    break;
                                                                }
                                                                else
                                                                    alarmClass = hmiUnified.AlarmClasses.Find(_settings.DefaultAlarmsClass);
                                                            }
                                                        }
                                                        break;
                                                    case "Bool":
                                                        var tagname = internalMember.Name.Replace(".", "_");
                                                        if(_settings.SimplifyTagname)
                                                            tagname = $"{globalDB.Name.Replace(_settings.BlockExtension, string.Empty)}_{tagname}";

                                                        HmiTag tag;
                                                        var plcTag = globalDB.Name.Contains(' ') ? $"\"{globalDB.Name}\"" : globalDB.Name;
                                                        foreach(var item in internalMember.Name.Split('.'))
                                                        {
                                                            if (item.Contains(" "))
                                                                plcTag += $".\"{item}\"";
                                                            else
                                                                plcTag += $".{item}";
                                                        }

                                                        

                                                        if (hmiUnified.Tags.FirstOrDefault(f => f.PlcTag == plcTag) is HmiTag hmiTag)
                                                        {
                                                            tag = hmiTag;
                                                            if(tag.Name != tagname)
                                                            {
                                                                var oldAlarm = hmiUnified.DiscreteAlarms.Find(tag.Name);
                                                                oldAlarm.Delete();

                                                                tag.Name = tagname;
                                                            }
                                                            exclusiveAccess.Text = $"{Extract.Core.Properties.Resources.BuildAlarms_UpdateAlarm} {globalDB.Name}/{exportTag.Name} {Extract.Core.Properties.Resources.BuildAlarms_AlarmOn} {device.Name}";
                                                        }
                                                        else
                                                        {
                                                            exclusiveAccess.Text = $"{Extract.Core.Properties.Resources.BuildAlarms_CreateAlarm} {globalDB.Name}/{exportTag.Name} {Extract.Core.Properties.Resources.BuildAlarms_AlarmOn} {device.Name}";

                                                            if (string.IsNullOrEmpty(folderName))
                                                                tag = hmiUnified.Tags.Create(tagname);
                                                            else
                                                            {
                                                                var tagTables = new List<HmiTagTable>();
                                                                tagTables.AddRange(hmiUnified.TagTables);

                                                                var tagTablesGroup = new List<HmiTagTableGroup>();
                                                                tagTablesGroup.AddRange(hmiUnified.TagTableGroups);
                                                                do
                                                                {
                                                                    if (tagTablesGroup.Count > 0)
                                                                    {
                                                                        tagTables.AddRange(tagTablesGroup[0].TagTables);
                                                                        tagTablesGroup.AddRange(tagTablesGroup[0].Groups);
                                                                        tagTablesGroup.RemoveAt(0);
                                                                    }
                                                                } while (tagTablesGroup.Count > 0);

                                                                if (!tagTables.Any(a => a.Name == folderName))
                                                                    hmiUnified.TagTables.Create(folderName);

                                                                tag = hmiUnified.Tags.Create(tagname, folderName);
                                                            }
                                                        }
                                                        tag.Connection = connexion.Name;
                                                        tag.PlcTag = plcTag;

                                                        var alarms = hmiUnified.DiscreteAlarms.Find(tagname) ?? hmiUnified.DiscreteAlarms.Create(tagname);
                                                        alarms.RaisedStateTag = tagname;
                                                        alarms.AlarmClass = alarmClass.Name;

                                                        alarms.Origin = globalDB.Name.Replace(_settings.BlockExtension, string.Empty);

                                                        foreach (var lang in activeLanguages)
                                                        {
                                                            if (alarms.EventText.Items.Find(lang) is MultilingualTextItem multilingualText)
                                                                multilingualText.Text = $"<body><p>{GetCommentText(lang.Culture, exportTag)}</p></body>";
                                                        }

                                                        tag = null;
                                                        plcTag = null;
                                                        alarms = null;
                                                        break;
                                                }
                                            }
                                        }
                                        _feedbackContext.Log(NotificationIcon.Success, $"{Extract.Core.Properties.Resources.BuildAlarms_UpdateOnDevice} {device.Name}");

                                        break;
                                }
                                softContainer = null;

                            }
                        }
                    }

                myFile.Close();
                File.Delete(path);

            }
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

        // Returns PlcSoftware or HmiSoftware
        private T GetDeviceSoftware<T>(Device device)
        {
            foreach (var deviceItem in device.DeviceItems)
            {
                return (T)Convert.ChangeType(deviceItem.GetService<SoftwareContainer>()?.Software, typeof(T));
            }
            return default;
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

        private SimaticML.SW.InterfaceSections.Member_T GetExportMember(Member member, SimaticML.SW.InterfaceSections.Member_T[] members)
        {
            IEnumerable<SimaticML.SW.InterfaceSections.Member_T> tempMembers = members;
            var tags = member.Name.Split('.');
            foreach (var tagComposition in tags.Take(tags.Length - 1))
            {
                tempMembers = tempMembers.FirstOrDefault(m => m.Name == tagComposition).Items.Where(i => i is SimaticML.SW.InterfaceSections.Member_T).Cast<SimaticML.SW.InterfaceSections.Member_T>();
            }

            return tempMembers.FirstOrDefault(m => m.Name == tags.Last());
        }

        private string GetCommentText(CultureInfo culture, SimaticML.SW.InterfaceSections.Member_T member)
        {
            string result = null;
            foreach (var item in member.Items.Where(i => i is SimaticML.SW.Common.Comment).Cast<SimaticML.SW.Common.Comment>())
            {
                result = item.MultiLanguageText.FirstOrDefault(comment => comment.Lang == culture.Name).Value;
            }
            return result;
        }
    }
}
