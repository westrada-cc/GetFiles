using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using  System.IO;

namespace GetFiles
{
    
    class Program
    {
      
        static void Main(string[] args)
        {
            string path = @"C:\Temp\ListOfTables.txt";
            int total = 0;
            List<string> fileName = new List<string>();
            StreamWriter myfile = new StreamWriter("ASPListFiles.txt");
            int count = 0;
           // string[] lstFiles = Directory.GetFiles(@"C:\globalWOLFASP\", "*.ASP");
            string[] lstFiles = Directory.GetFiles(@"C:\\ASPModulesTemp\", "*.ASP");
            using (StreamWriter listOfTablesInFiles = File.CreateText(path))
            {
                listOfTablesInFiles.WriteLine("List Of Tables....\n");
                listOfTablesInFiles.WriteLine("\n");
            }

            foreach (string file in lstFiles)
            {
                Console.WriteLine(GetJustFileName(file));
                myfile.Write(GetJustFileName(file));
                fileName.Add(GetJustFileName(file));
                total += ReadFileContent(file);
               //ReadFileASP(file);

                myfile.Write("\n");
                count++;
            }

            //using (StreamWriter listOfTablesInFiles = File.CreateText(path))
            //{
            //    listOfTablesInFiles.WriteLine(GetJustFileName(file));
            //    listOfTablesInFiles.WriteLine("\n");
            //}

            //foreach (string s in fileName)
            //{
            //    Console.WriteLine(s);
            //}
            myfile.Close();
            //Console.WriteLine("\n");
            //Console.WriteLine("Total files checked: {0} : Total Files with INSERT INTO {1}", count, total);
            //removeList();
            Console.WriteLine("Total Files Read {0}", count);
            Console.WriteLine("=================   DONE!  =============================");
            Console.WriteLine("=================   DONE!! =============================");
            Console.ReadKey();
        }


        static private void removeList()
        {
            char[] charsToTrim = {'>'};
            string longWod = "apanle>";
           int start = 3;
            List<string> test = new List<string>();
            test.Add("Horse");
            test.Add("Cat");
            test.Add("Bird");
            test.Add("Bonny");
            test.Add("Bat");
            foreach(string s in test)
            {
                Console.WriteLine(s);
            }
            test.RemoveRange(start, 2);
            foreach(string s in test)
            {
                Console.WriteLine(s);
            }
            string last = longWod.TrimEnd(charsToTrim);
            Console.WriteLine(last);
            Console.ReadLine();
        }


        /// <summary>
        /// slipt the string to get just the file name
        /// </summary>
        /// <param name="name">takes the path where the file is</param>
        /// <returns>returns the file's name</returns>
        static private string GetJustFileName(string name)
        {
            string[] fileName = name.Split('\\');
            string newfile = fileName[3];
            return newfile;
        }

        static private void ReadFileASP(string file)
        {
            List<string> fileContent = new List<string>();
            StreamReader filReader = new StreamReader(file);
            string toBeReplace = string.Empty;
           // string newString = "<%\n\tQS_GetPropertyType_ControlClassName = \n\tsuccess = QS_GetPropertyType()\n %> ";
            string className = string.Empty;
            string id = "id=";
            string isInfile = @"""PropertyType""";
            string endOfString = "</select>";
            string fullText = id + isInfile;
            string line = string.Empty;
            int start = 0;
            int end = 0;
            int x = 1;
            int result = 0;
            int cutoff = 0;
            int stop = 0;
            while ((line = filReader.ReadLine()) != null)
            {
                fileContent.Add(line);
               // counter++;
            }
            filReader.Close();
            foreach(string s in fileContent)
            {
                
                if (s.Contains(fullText))
                {
                    char[] charToTrim = { '>' };
                    string[] fileName = s.Split('=');
                    string newfile = fileName[5];
                    className = newfile.TrimEnd(charToTrim);
                    //Console.WriteLine("\n");
                    //Console.WriteLine(s);
                    start = x;
                    stop = start;
                }
                if(s.Contains(endOfString))
                {
                    end = x;
                    toBeReplace += s;
                    cutoff = end;
                }
                //if (start != end)
                //{
                    
                //}
                if((start == x) && (start != end))
                {
                    toBeReplace += s;
                    start++;
                }
                if((start == x) && (end == x))
                {
                    string[] input = {
                                 "\t\t\t\t\t\t\t\t\t<%",
                                 "\t\t\t\t\t\t\t\t\t\tQS_GetPropertyType_ControlClassName = " + className,
                                 "\t\t\t\t\t\t\t\t\t\tsuccess = QS_GetPropertyType()",
                                 "\t\t\t\t\t\t\t\t\t%>"
                             };
                    result = cutoff - stop;
                   //using (StreamWriter wrString = new StreamWriter("test.asp"))
                   // {
                   //     string newtext = string.Empty;
                   //     newtext.Replace(toBeReplace, newString);
                   //     wrString.Write(newtext);
                   // }
                   // break;
                    string newfileName = GetJustFileName(file);
                    fileContent.RemoveRange(stop - 1, result+1);
                    fileContent.InsertRange(stop - 1, input);
                    File.WriteAllLines(newfileName, fileContent);
                    Console.WriteLine("the File {0}, was changed successful!", file);
                    s.DefaultIfEmpty();
                    break;
                }
                x++;
            }


        }



        /// <summary>
        /// put the content of the file into a list
        /// </summary>
        /// <param name="file"></param>
        static private int ReadFileContent(string file)
        {
            string path = @"C:\Temp\ListOfTables.txt";
            List<string> tableName = new List<string>();
            tableName.Add("City");
            tableName.Add("CompanyListings");
            tableName.Add("CompanyPropertyType");
            tableName.Add("ConCategory");
            tableName.Add("ConCompany");
            tableName.Add("Contacts");
            tableName.Add("ContactActionRequired");
            tableName.Add("ContactActionTaken");
            tableName.Add("ContactCategory");
            tableName.Add("Drip_Campaign");
            tableName.Add("Drip_Letters");
            tableName.Add("Event");
            tableName.Add("InfoTopic");
            tableName.Add("InfoTopicDynamic");
            tableName.Add("LC_Lead");
            tableName.Add("LC_LeadCategory");
            tableName.Add("LC_OfficeHours");
            tableName.Add("LC_Rule");
            tableName.Add("LC_SOB");
            tableName.Add("LeadEmailTemplateSettings");
            tableName.Add("LoginLinks");
            tableName.Add("MasterWebFormQuestions");
            tableName.Add("Message");
            tableName.Add("MessageProvider");
            tableName.Add("movingWOLF_Campaign");
            tableName.Add("movingWOLF_Letters");
            tableName.Add("SEO_Pages");
            tableName.Add("SearchDefinition");
            tableName.Add("ShowingAgentType");
            tableName.Add("ShowingCallStatus");
            tableName.Add("ShowingEmailTemplateSettings");
            tableName.Add("ShowingExceptions");
            tableName.Add("ShowingType");
            tableName.Add("SocialMediaUse");
            tableName.Add("UserGroup");
            tableName.Add("WebForm");
            tableName.Add("WIGO_Category");
            tableName.Add("xSettings");
            string isInfile = "SELECT ";
            string isDelete = "DELETE ";
            string isFrom = "FROM";
            string isUpdate = "UPDATE";
            string isJoin = "JOIN";
            int counter = 0;
            string strInFile = string.Empty;
            List<string> fileContent = new List<string>();
            StreamReader filReader = new StreamReader(file);
            string line = string.Empty;
            string listOfTables = string.Empty;
            int x = 1;
            int nFile = 1;
            int fileFound = 0;
            while ((line = filReader.ReadLine()) != null)
            {
                fileContent.Add(line);
                counter++;
                
            }

            filReader.Close();
            foreach (string s in fileContent)
            {
                foreach (string s1 in tableName)
                {
                    //strInFile = isInfile + s1;
                    if (((s.Contains(isInfile)) && (s.Contains(s1))) || ((s.Contains(isDelete)) && (s.Contains(s1))) || ((s.Contains(isFrom)) && (s.Contains(s1))) || ((s.Contains(isUpdate) && (s.Contains(s1)))) || ((s.Contains(isJoin) && s.Contains(s1))) || (s.Contains(s1)))
                    {
                        
                        if (nFile == 1)
                        {
                            Console.WriteLine("\n");
                            Console.WriteLine(file);
                            using ( StreamWriter listOfTablesInFiles = File.AppendText(path))
                            {
                                listOfTablesInFiles.WriteLine("File Name: \t " + GetJustFileName(file));
                                listOfTablesInFiles.WriteLine("\n");
                            }
                            
                           // Console.WriteLine("\n");
                            fileFound++;
                        }
                        
                        Console.WriteLine("Table {0} found at line: {1}", s1, x);
                        listOfTables = "Table " + s1 + " found at line " + x;
                        using(StreamWriter listOfTablesInFiles = File.AppendText(path))
                        {
                            listOfTablesInFiles.WriteLine(listOfTables);
                            listOfTablesInFiles.Write("\n");
                        }
                        //listOfTablesInFiles.Append(listOfTables);
                        listOfTables.DefaultIfEmpty();
                        
                        //Console.ReadKey();
                        nFile++;
                    }
                }
               
                x++;
            }

            //listOfTablesInFiles.Close();
            return fileFound;

        }
    }
}
