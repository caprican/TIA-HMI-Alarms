using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HmiUnified.HmiConnections;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;

namespace Extract.Models
{
    public class ProjectConnexion
    {
        public HmiSoftware Hmi { get; }
        public HmiConnection Connexion { get; }
        
        public PlcSoftware Plc { get; }
        public PlcBlock PlcBlock { get; }


        public ProjectConnexion(HmiSoftware hmi, PlcSoftware plc, HmiConnection connexion)
        {
            Hmi = hmi;
            Plc = plc;
            Connexion = connexion;
        }

        public ProjectConnexion(HmiSoftware hmi, PlcBlock plcBlock, HmiConnection connexion)
        {
            Hmi = hmi;
            Connexion = connexion;
            PlcBlock = plcBlock;
        }
    }
}
