using System.Collections.Generic;

namespace ClassLibrary
{
    internal class Armature
    {
        public string name;
        public string pathB1;
        public string pathB2;
        public List<string> values;
        public List<int> valuesColumn;        

        public Armature(string name, List<string> values, List<int> valuesColumn, string directoryPath)
        {
            this.name = name;
            this.values = values;
            this.valuesColumn = valuesColumn;
            pathB1 = directoryPath + name + "_B1";
            pathB2 = directoryPath + name + "_B2";
        }
    }
}
