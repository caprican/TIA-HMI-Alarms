using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
using SimaticML;
using SimaticML.SW;
using SimaticMLBlock = SimaticML.SW.Blocks;
using System.Text.RegularExpressions;

namespace TIA_Extract
{
    public class SKFAlarmsAddIn : ContextMenuAddIn
    {
        /// <summary>
        /// The instance of TIA Portal in which the Add-In works.
        /// <para>Enables Add-In to interact with TIA Portal.</para>
        /// </summary>
        private readonly TiaPortal _tiaPortal;

        /// <summary>
        /// Opened TIA Portal project
        /// </summary>
        internal Project project;

        /// <summary>
        /// The display name of the Add-In.
        /// </summary>
        private const string s_DisplayNameOfAddIn = "SKF Alarms HMI";

        private FeedbackContext _feedbackContext;

        /// <summary>
        /// The constructor of the AddIn.
        /// Creates an instance of the class AddIn.
        /// <para>- Called from AddInProvider when the first right-click is performed in TIA Portal.</para>
        /// <para>- The base class' constructor of AddIn will also be executed.</para>
        /// </summary>
        /// <param name="tiaPortal">
        /// Represents the instance of TIA Portal in which the Add-In will work.
        /// </param>
        public SKFAlarmsAddIn(TiaPortal tiaPortal) : base(s_DisplayNameOfAddIn)
        {
            _tiaPortal = tiaPortal;
            try
            {
                _feedbackContext = _tiaPortal.GetFeedbackContext();
            }catch { }
            
            //#if DEBUG
            //            Core.Properties.Resources.Culture = new CultureInfo("en-US");
            //#else
            //            TiaPortalSettingsFolder generalSettingsFolder = _tiaPortal.SettingsFolders.Find("General");
            //            TiaPortalSetting UILanguageSetting = generalSettingsFolder.Settings.Find("UserInterfaceLanguage");
            //            Core.Properties.Resources.Culture = UILanguageSetting.Value as CultureInfo;
            //#endif
            project = _tiaPortal.Projects.First(x => x.IsPrimary);

            // Gets the TiaPortal version from the project filename extension
            var projectVersion = project.Path.Extension;
            var _tiaVersion = projectVersion.Substring(3).Replace("_", ".");

            var dirPath = Path.Combine(Environment.ExpandEnvironmentVariables("%TEMP%"), AppDomain.CurrentDomain.FriendlyName);
            var dir = Directory.CreateDirectory(dirPath);
            var _tempFilePath = dir.FullName;

            var path = project.Path.Directory.FullName;
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
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>(Properties.Resource.ContextMenu_GlobalDb, OnGenerateClick,
                (menuSelectionProvider) => menuSelectionProvider.GetSelection().Any(engineeringObject => engineeringObject is SimaticSW.GlobalDB) ? MenuStatus.Enabled : MenuStatus.Disabled);

            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>(Properties.Resource.ContextMenu_UserGroup, OnGenerateClick,
                (menuSelectionProvider) => menuSelectionProvider.GetSelection().Any(engineeringObject => engineeringObject is SimaticSW.PlcBlockUserGroup) ? MenuStatus.Enabled : MenuStatus.Disabled);

            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>(Properties.Resource.ContextMenu_HmiTag, OnGenerateClick,
                (menuSelectionProvider) => menuSelectionProvider.GetSelection().Any(engineeringObject => engineeringObject is TagTable) ? MenuStatus.Enabled : MenuStatus.Disabled);
        }

        private void OnGenerateClick(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            if(_feedbackContext is null)
            {
                try
                {
                    _feedbackContext = _tiaPortal.GetFeedbackContext();
                }
                catch { }
            }
            using (var exclusiveAccess = _tiaPortal.ExclusiveAccess())
            {
                foreach (var plcBlock in menuSelectionProvider.GetSelection<IEngineeringObject>())
                {
                    if(exclusiveAccess.IsCancellationRequested)
                    {
                        _feedbackContext.Log(NotificationIcon.Information, Properties.Resource.CancelByUser);
                        return;
                    }

                    switch(plcBlock)
                    {
                        case SimaticSW.GlobalDB globalDb:
                            BuildAlarms(exclusiveAccess, globalDb);
                            break;
                        case SimaticSW.PlcBlockUserGroup blockGroup:

                            List<SimaticSW.PlcBlockUserGroup> tempGroupList = new List<SimaticSW.PlcBlockUserGroup> { blockGroup };
                            var asUpdateAlarms = false;
                            do
                            {
                                if (tempGroupList[0].Blocks.Any(block => block is SimaticSW.GlobalDB))
                                {
                                    foreach (var plcDataBlock in blockGroup.Blocks.Where(bloc => bloc is SimaticSW.GlobalDB).Cast<SimaticSW.GlobalDB>())
                                    {
                                        asUpdateAlarms = true;
                                        BuildAlarms(exclusiveAccess, plcDataBlock);
                                        _feedbackContext.Log(NotificationIcon.Success, "Alarmes mis à jours");
                                    }
                                }
                                if (tempGroupList[0].Groups.Count > 0)
                                {
                                    tempGroupList.AddRange(tempGroupList[0].Groups);
                                }

                                tempGroupList.RemoveAt(0);
                            } while (tempGroupList.Count > 0);

                            if(!asUpdateAlarms)
                            {
                                _feedbackContext.Log(NotificationIcon.Information, $"{Properties.Resource.EmptyGlobalDb} {blockGroup.Name}");
                                break;
                            }

                            break;
                        case TagTable hmiTagTable:
                            var hmiSoft = GetDeviceSoftware<HmiSoftware>(hmiTagTable);

                            foreach (var connection in hmiSoft.Connections)
                            {
                                var devices = new List<Device>();
                                devices.AddRange(project.Devices);

                                var deviceGroup = new List<DeviceUserGroup>();
                                deviceGroup.AddRange(project.DeviceGroups);
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
                                        foreach (var plcGroupBlock in group.Blocks.Where(bloc => bloc.Name.EndsWith("_Defauts")))
                                        {
                                            if (plcGroupBlock is SimaticSW.GlobalDB dB)
                                            {
                                                BuildAlarms(exclusiveAccess, dB);
                                                _feedbackContext.Log(NotificationIcon.Success, "Alarmes mis à jours");
                                            }

                                        }
                                    }
                                }
                            }
                            break;
                    }
                }
                _feedbackContext.Log(NotificationIcon.Success, Properties.Resource.BuildAlarmsEnded);
            }

        }

        private void BuildAlarms(ExclusiveAccess exclusiveAccess, SimaticSW.GlobalDB globalDB)
        {
            _feedbackContext.Log(NotificationIcon.Information, $"Extraction des alarmes depuis {globalDB.Name}");

            exclusiveAccess.Text = $"Génération des messages depuis {globalDB.Name}";

            if (globalDB.Name.EndsWith("_Defauts"))
            {
                var folderName = string.Empty;
                if (globalDB.Parent is SimaticSW.PlcBlockUserGroup group)
                {
                    folderName = group.Name;
                }

                var plcSoft = GetDeviceSoftware<PlcSoftware>(globalDB);

                if (!globalDB.IsConsistent)
                {
                    var comp = globalDB.GetService<ICompilable>();
                    if (comp.Compile().State != CompilerResultState.Success)
                    {
                        _feedbackContext.Log(NotificationIcon.Error, $"{globalDB.Name} n'est pas compilable.");
                        return;
                    }
                }

                var path = Path.Combine(project.Path.Directory.FullName, "UserFiles", globalDB.Name + ".xml");

                try
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                    globalDB.Export(new FileInfo(path), ExportOptions.None);
                }
                catch (Exception ex)
                {
                    _feedbackContext.Log(NotificationIcon.Error, ex.Message);
                    return;
                }

                var serializer = new XmlSerializer(typeof(SimaticML.Document));
                var myFile = new FileStream(path, FileMode.Open);
                if (serializer.Deserialize(myFile) is Document document && document.SWBlocks is SimaticMLBlock.GlobalDB Db)
                {
                    var devices = new List<Device>();
                    devices.AddRange(project.Devices);

                    var deviceGroup = new List<DeviceUserGroup>();
                    deviceGroup.AddRange(project.DeviceGroups);
                    do
                    {
                        if(deviceGroup.Count > 0)
                        {
                            devices.AddRange(deviceGroup[0].Devices);
                            deviceGroup.AddRange(deviceGroup[0].Groups);
                        }
                        deviceGroup.RemoveAt(0);
                    } while (deviceGroup.Count > 0);


                    foreach (var device in devices)
                    {
                        foreach (var deviceItem in device.DeviceItems)
                        {
                            if (exclusiveAccess.IsCancellationRequested)
                            {
                                _feedbackContext.Log(NotificationIcon.Information, "La tâche à été annulé par l'utilisateur");
                                return;
                            }

                            var exportDb = Db.Attributes.Interface.Sections[0];

                            var softContainer = deviceItem.GetService<SoftwareContainer>();
                            switch (softContainer?.Software)
                            {
                                case HmiTarget hmi:
                                    _feedbackContext.Log(NotificationIcon.Information, $"Générateur non compatible avec {device.Name}.");
                                    break;
                                case HmiSoftware hmiUnified:
                                    var connexion = hmiUnified.Connections.FirstOrDefault(con => con.Partner == plcSoft.Name);

                                    if (connexion is null)
                                    {
                                        _feedbackContext.Log(NotificationIcon.Information, $"Pas de connexion trouvée entre {device.Name} et {plcSoft.Name}.");
                                        break;
                                    }

                                    var alarmClass = hmiUnified.AlarmClasses.Find("Alarm");

                                    foreach (var internalMember in globalDB.Interface.Members)
                                    {
                                        if (exclusiveAccess.IsCancellationRequested)
                                        {
                                            _feedbackContext.Log(NotificationIcon.Information, Properties.Resource.CancelByUser);
                                            return;
                                        }

                                        var exportTag = new SimaticML.SW.Member { Members = exportDb.Members.ToList() };
                                        foreach (var tagComposition in internalMember.Name.Split('.'))
                                        {
                                            exportTag = exportTag.Members.FirstOrDefault(f => f.Name == tagComposition);
                                        }

                                        if (exportTag != null)
                                        {
                                            switch(exportTag.Datatype)
                                            {
                                                case "Struct":
                                                    foreach (var lang in project.LanguageSettings.ActiveLanguages)
                                                    {
                                                        var txt = exportTag.Comment.Find(comment => comment.Lang == lang.Culture.Name).Value;
                                                        if(!string.IsNullOrEmpty(txt))
                                                        {
                                                            var match = Regex.Match(txt, @"(.+?)\[AlarmClass=""[^""]*""](.+?)", RegexOptions.IgnoreCase);
                                                            if(match.Success && hmiUnified.AlarmClasses.Find(match.Groups[1].Value) is HmiAlarmClass groupAlarmClass)
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
                                                            {
                                                                alarmClass = hmiUnified.AlarmClasses.Find("Alarm");
                                                            }
                                                        }
                                                    }
                                                    break;
                                                case "Bool":
                                                    var tagname = internalMember.Name.Replace(".", "_");
                                                    tagname = $"{globalDB.Name.Replace("_Defauts", string.Empty)}_{tagname}";

                                                    HmiTag tag;
                                                    if (hmiUnified.Tags.Find(tagname) is HmiTag hmiTag)
                                                    {
                                                        tag = hmiTag;
                                                        exclusiveAccess.Text = $"Mise à jour de l'alarme depuis {globalDB.Name}/{exportTag.Name} dans {device.Name}";
                                                    }
                                                    else
                                                    {
                                                        exclusiveAccess.Text = $"Création de l'alarme depuis {globalDB.Name}/{exportTag.Name} dans {device.Name}";
                                                        tag =  (string.IsNullOrEmpty(folderName) ? hmiUnified.Tags.Create(tagname) : hmiUnified.Tags.Create(tagname, folderName));

                                                    }
                                                    tag.Connection = connexion.Name;
                                                    tag.PlcTag = $"{globalDB.Name}.{internalMember.Name}";
                                                    

                                                    var alarms = hmiUnified.DiscreteAlarms.Find(tagname) ?? hmiUnified.DiscreteAlarms.Create(tagname);
                                                    alarms.RaisedStateTag = tagname;
                                                    alarms.AlarmClass = alarmClass.Name;
                                                    alarms.Origin

                                                    foreach (var lang in project.LanguageSettings.ActiveLanguages)
                                                    {
                                                        if (alarms.EventText.Items.Find(lang) is MultilingualTextItem multilingualText)
                                                        {
                                                            var txt = exportTag.Comment.Find(comment => comment.Lang == lang.Culture.Name).Value;
                                                            multilingualText.Text = $"<body><p>{txt}</p></body>";
                                                        }
                                                    }
                                                    break;
                                            }
                                        }
                                    }
                                    _feedbackContext.Log(NotificationIcon.Success, $"Alarmes mis à jours sur {device.Name}");

                                    break;
                            }
                        }
                    }
                }

                myFile.Close();
                File.Delete(path);

            }
        }

        // Returns PlcSoftware or HmiSoftware
        private T GetDeviceSoftware<T>(Device device)
        {
            foreach (var deviceItem in device.DeviceItems)
            {
                return (T) Convert.ChangeType(deviceItem.GetService<SoftwareContainer>()?.Software, typeof(T));
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
            return plc is null ? default : (T) Convert.ChangeType(plc, typeof(T));
        }
    }
}
