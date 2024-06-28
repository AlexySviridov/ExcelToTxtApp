using System;
using System.Collections.Generic;
using System.Linq;

namespace ClassLibrary
{
    internal class Armature
    {
        public string Name;
        public string PathB1;
        public string PathB2;
        public List<string> Values;
        public List<int> ValuesColumn;        

        public Armature(string name, List<string> values, List<int> valuesColumn, string directoryPath)
        {
            Name = name;
            Values = values;
            ValuesColumn = valuesColumn;
            PathB1 = directoryPath + name + "_B1";
            PathB2 = directoryPath + name + "_B2";
        }

        public void IdentifiedArmatureType(string[] bansArray, string[] commandsArray, out TypeArmature typeArmature)
        {
            var bansInFirstField = false;
            var commandsInFirstField = false;
            var commandsInSecondField = false;

            foreach (string value in Values)
            {
                var firstFieldValue = value.Split('/')[0];
                if (value.Split('/').Length == 1)
                {
                    if (bansArray.Contains(firstFieldValue)) bansInFirstField = true;
                    else if (commandsArray.Contains(firstFieldValue)) commandsInFirstField = true;
                    else throw new Exception("Неопознанный запрет или команда: " + firstFieldValue);
                }
                else if (value.Split('/').Length > 1)
                {
                    var secondFieldValue = value.Split('/')[1];
                    if (bansArray.Contains(firstFieldValue)) bansInFirstField = true;
                    if (commandsArray.Contains(secondFieldValue)) commandsInSecondField = true;
                    else if (secondFieldValue != "Руч") throw new Exception("Неопознанный запрет или команда: " + secondFieldValue);
                }
            }

            if (!bansInFirstField) typeArmature = TypeArmature.BansNotExists;
            else if (!commandsInFirstField && !commandsInSecondField) typeArmature = TypeArmature.CommandsNotExist;
            else if (bansInFirstField && commandsInFirstField) typeArmature = TypeArmature.BansAndCommandsExistInFirstField;
            else if (bansInFirstField && commandsInSecondField) typeArmature = TypeArmature.CommandsExistInSecondField;
            else throw new Exception("Обработать логику данной арматуры (" + Name + ") не представляется возможным для текущей версии программы O_o");
        }
    }
}
