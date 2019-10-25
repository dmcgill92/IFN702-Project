using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using RandomNameGeneratorLibrary;

namespace Random_Data_Generator
{
	class DataGenerator
	{
		List<string> firstNames = new List<string>();
		List<string> lastNames = new List<string>();
		List<string> studentNums = new List<string>();
		List<int> results = new List<int>();
		int num = 100;

		static void Main(string[] args)
		{
			DataGenerator myGenerator = new DataGenerator();
			int num = myGenerator.num;
			myGenerator.GenerateNames(num);
			myGenerator.GenerateStudentNumbers(num);
			myGenerator.GenerateResults(num);
			myGenerator.CreateExcel("Results", true);
			myGenerator.CreateExcel("Students", false);
		}

		public void GenerateNames(int num)
		{
			PersonNameGenerator personGenerator = new PersonNameGenerator();
			Random random = new Random();

			for (int iter = 0; iter < num; iter++)
			{
				if (random.Next(2) == 1)
				{
					firstNames.Add(personGenerator.GenerateRandomMaleFirstName());
				}
				else
				{
					firstNames.Add(personGenerator.GenerateRandomFemaleFirstName());
				}
				lastNames.Add(personGenerator.GenerateRandomLastName());
			}
		}

		public void GenerateStudentNumbers(int num)
		{
			Random random = new Random();

			for (int iter = 0; iter < num; iter++)
			{
				studentNums.Add(random.Next(900000, 1400000).ToString());
			}
		}

		public void GenerateResults(int num)
		{
			Random random = new Random();

			for (int iter = 0; iter < num; iter++)
			{
				results.Add(random.Next(10, 100));
			}
		}

		public void CreateExcel(string path, bool isDirty)
		{
			Application xlApp;
			Workbook workBook;
			Worksheet workSheet;

			xlApp = new Application();
			if (xlApp == null)
			{
				Console.WriteLine("Excel is not properly installed!!");
				return;
			}

			xlApp.Visible = false;
			xlApp.DisplayAlerts = false;

			workBook = xlApp.Workbooks.Add(Type.Missing);
			workSheet = (Worksheet)workBook.ActiveSheet;
			if (isDirty)
            { 
				workSheet.Name = "Results";
                workSheet.Cells[1, 1] = "Student number";
                workSheet.Cells[1, 2] = "Surname";
                workSheet.Cells[1, 3] = "Initial";
                workSheet.Cells[1, 4] = "Score";
            }
            else
            { 
				workSheet.Name = "Students";
                workSheet.Cells[1, 1] = "Last Name";
                workSheet.Cells[1, 2] = "First Name";
                workSheet.Cells[1, 3] = "Username";
				workSheet.Cells[1, 4] = "Assessment 1";
				workSheet.Cells[1, 5] = "Assessment 2";
				workSheet.Cells[1, 6] = "Exam";
            }

            Random rand = new Random();

			for (int iter = 0; iter < firstNames.Count; iter++)
			{
                if(isDirty)
                {
					Range range = workSheet.Range["A1", "C" + num];
					range.NumberFormat = "@";
					if (rand.Next(0, 10) == 0)
                    {
						workSheet.Cells[iter + 2, 1] = DistortData(studentNums[iter], false);
                    }
                    else
                    {
				        workSheet.Cells[iter + 2, 1] = studentNums[iter];
                    }

                    if(rand.Next(0, 10) == 0)
                    {
                        workSheet.Cells[iter + 2, 2] = DistortData(lastNames[iter], true).ToUpper();
                    }
                    else
                    {
				        workSheet.Cells[iter + 2, 2] = lastNames[iter].ToUpper();
                    }

                    if (rand.Next(0, 10) == 0)
                    {
                        workSheet.Cells[iter + 2, 3] = DistortData(firstNames[iter][0].ToString(), true).ToUpper();
                    }
                    else
                    {
                        workSheet.Cells[iter + 2, 3] = firstNames[iter][0].ToString().ToUpper();
                    }

					workSheet.Cells[iter + 2, 4] = results[iter];
                }
                else
                {
					Range range = workSheet.Range["A1", "C" + num];
					range.NumberFormat = "@";
                    workSheet.Cells[iter + 2, 1] = lastNames[iter];
                    workSheet.Cells[iter + 2, 2] = firstNames[iter];
                    workSheet.Cells[iter + 2, 3] = "n" + studentNums[iter];
                    workSheet.Cells[iter + 2, 4] = "";
                }
			}

			workBook.SaveAs(path);
			if(isDirty)
			{
				workBook.SaveAs(path, XlFileFormat.xlCSV);
			}

			workBook.Close();
			xlApp.Quit();
		}

		public string DistortData(string word, bool isName)
		{
			Random rand = new Random();
			char[] temp = word.ToCharArray();

			if(rand.Next(0, 10) == 1)
			{
				for( int j = 0; j < temp.Length; j++)
				{
					temp[j] = ' ';
				}
			}
			else
			{
				for( int i = 0; i < word.Length; i++)
				{
					int choice = rand.Next(0, 10);
					if (choice == 1)
					{
						if(isName)
						{
                            int num = rand.Next(0, 26);
							temp[i] = (char)('a' + num);
						}
						else
						{
							string num = rand.Next(0, 9).ToString();
							temp[i] = num[0];
						}
					}
					else
					if (choice >= 2 && choice <= 5)
					{
						temp[i] = ' ';
					}
				}
			}

			string tempStr = new string(temp);
			return tempStr;
		}
	}
}

