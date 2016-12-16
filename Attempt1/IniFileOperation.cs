using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace Attempt1
{
    class IniFileOperation
    {
        private string iniFilePath;
        public string IniFilePath
        {
            get;
            set;
        }

        [DllImport("kernel32.dll")]
        // section：INI文件中的段落；key：INI文件中的关键字；
        // val：INI文件中关键字的数值；filePath：INI文件的完整的路径和名称。
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32.dll")]
        // section：INI文件中的段落名称；key：INI文件中的关键字；
        // def：无法读取时候时候的缺省数值；retVal：读取数值；size：数值的大小；filePath：INI文件的完整路径和名称。
        private static extern int GetPrivateProfileString(string section, string key, string def,
            StringBuilder retVal, int size, string filePath);

        public IniFileOperation(string iniFilePath)
        {
            this.iniFilePath = iniFilePath;
        }

        public void IniWriteValue(string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value, this.iniFilePath);
        }

        public string IniReadValue(string Section, string Key)
        {
            StringBuilder strTemp = new StringBuilder(500);
            int i = GetPrivateProfileString(Section, Key, "", strTemp, 500, this.iniFilePath);
            return strTemp.ToString();
        }

        public bool ExistINIFile()
        {
            return File.Exists(this.iniFilePath);
        }
    }
}
