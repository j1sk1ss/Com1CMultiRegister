using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using COMAdmin;
using Microsoft.Win32;

namespace Com1CMultiRegister {
    public class Manager {
        public Manager(string pattern) {
            Path    = findPath(""); // Needs name of 1C application
            Pattern = pattern;
        }

        public Manager(string path, string pattern) {
            Path    = path;
            Pattern = pattern;
        }


        private string Path { get; set; }
        private string Pattern { get; set; }


        // Function get list off all applications and find 1C
        // @return path to 1C application
        private string findPath(string programName) {
            var progReturn = new ManagementObjectSearcher("Select * from Win32_Product").Get();
            foreach (var prog in progReturn) {
                if (prog["Name"].ToString() != programName) continue;
                return prog["InstallLocation"].ToString();
            }

            return null;
        }

        // Function get components
        // @return List of components
        public List<Component> GetComponents() {
            if (!Directory.Exists(Path)) return null;

            var components = new List<Component>();
            var catalogs   = Directory.GetDirectories(Path, Pattern);
            foreach (var path in catalogs) {
                var verName = path.Split('\\').Last();
                var comPath = path + @"\bin\comcntr.dll";
                if (!File.Exists(comPath)) continue;

                components.Add(new Component(comPath, verName));
            }

            return components;
        }

        public void CreateComponents(List<Component> components) {
            ICatalogObject catalogObject = null;
            ICatalogObject mainConnector = null;

            COMAdminCatalog catalog = new COMAdminCatalog();
            ICatalogCollection appCollection = (ICatalogCollection)catalog.GetCollection("Applications");
            appCollection.Populate();
            
            // Try to find existed catalog
            foreach (ICatalogObject obj in appCollection) {
                if (catalogObject.Name.ToString().ToLower() == "1cv8") {
                    catalogObject = obj;
                    break;
                }
            }

            // If catalog not found creates new catalog
            if (catalogObject == null) {
                ICatalogObject new1Cv8App = (ICatalogObject)appCollection.Add();
                new1Cv8App.Value["Name"] = "1cv8";
                new1Cv8App.Value["Activation"] = COMAdminActivationOptions.COMAdminActivationInproc;
                appCollection.SaveChanges();

                catalogObject = new1Cv8App;
            }

            ICatalogCollection comCollections = (ICatalogCollection)appCollection.GetCollection("Components", catalogObject.Key);
            comCollections.Populate();

            foreach (ICatalogObject obj in comCollections) {
                if (obj.Name.ToString() == "V83.COMConnector.1")
                    mainConnector = obj;

                var arrayName = obj.Name.ToString().Split('_');
                var ver       = arrayName[arrayName.Length - 1];
                var findVer   = components.Where(s => s.Version == ver);

                if (findVer.Count() != 0)
                    components.Remove(findVer.First());
            }

            if (mainConnector == null) {
                catalog.InstallComponent("1cv8", components.Last().Path, "", "");
                mainConnector = findMainCom(comCollections);
            }

            foreach (var itemVer in components) 
                catalog.AliasComponent("1cv8", mainConnector.get_Value("CLSID").ToString(), "", "V83.COMConnector_" + itemVer.Version, "");

            comCollections.Populate();

            foreach (ICatalogObject obj in comCollections) {
                if (obj.Name.ToString() == "V83.COMConnector.1")
                    continue;
               
                var arrayName = obj.Name.ToString().Split('_');
                Component findingItem = null;

                try {
                    findingItem = components.First(s => s.Version == arrayName[arrayName.Length - 1]);
                }
                catch { continue; }

                if (findingItem == null) continue;

                var clsid   = obj.get_Value("CLSID").ToString();
                var readKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Classes\\Wow6432Node\\CLSID\\" + clsid + "\\InprocServer32", true);
                readKey.SetValue("", findingItem.Path);
                readKey.Close();
            }

            comCollections.SaveChanges();
        }

        private ICatalogObject findMainCom(ICatalogCollection comCollections) {
            comCollections.Populate();
            foreach (ICatalogObject catalogObject in comCollections) 
                if (catalogObject.Name.ToString() == "V83.COMConnector.1") 
                    return catalogObject;

            return null;
        }
    }
}
