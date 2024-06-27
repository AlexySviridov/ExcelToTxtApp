using System.Collections.Generic;

namespace ClassLibrary
{
    internal class Armature
    {
        public string name;
        public List<string> values;
        public List<int> valuesColumn;

        public Armature(string name, List<string> values, List<int> valuesColumn)
        {
            this.name = name;
            this.values = values;
            this.valuesColumn = valuesColumn;
        }
    }
}
