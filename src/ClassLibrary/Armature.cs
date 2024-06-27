using System.Collections.Generic;

namespace ClassLibrary
{
    internal class Armature
    {
        public string name;
        public int numberRow;
        public List<string> values;

        public Armature(string name, int numberRow, List<string> values)
        {
            this.name = name;
            this.numberRow = numberRow;
            this.values = values;
        }
    }
}
