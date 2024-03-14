namespace Com1CMultiRegister {
    public class Component {
        public Component(string path, string version) {
            Path    = path;
            Version = version;
        }


        public string Path { get; set; }
        public string Version { get; set; }
        public string Clsid { get; set; }
    }
}
