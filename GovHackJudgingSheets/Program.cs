﻿using System;
using System.Collections.Generic;
using System.IO;
using CsvHelper;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.DataValidation;

namespace GovHackJudgingSheets
{
    class MainClass
    {
        public static Dictionary<string, Project> projects;
        public static Dictionary<string, Award> awards;

        public static void Main(string[] args)
        {
            Console.WriteLine("ReadInputs()");
            ReadInputs();
            Console.WriteLine("CreateJudgingSheets()");
            CreateJudgingSheets();
            Console.WriteLine("ReadJudgedSheets()");
            ReadJudgedSheets();
            Console.WriteLine("OutputJudgedResults()");
            OutputJudgedResults();
            Console.WriteLine("BaselineJudges()");
            BaselineJudges();

        }

        public static void ReadInputs() {
            ReadProjects();
            ReadHackerspaceVideoUhh();
            ReadAwards();
        }

        public static void ReadProjects() {
            projects = new Dictionary<string, Project>();

			/**
             * https://2017.hackerspace.govhack.org/projects.csv
             */

			using (StreamReader srProjects = new StreamReader("Input/hackerspace-projects.csv"))
			{
				using (CsvReader csvProjects = new CsvReader(srProjects))
				{
					csvProjects.ReadHeader();
					while (csvProjects.Read())
					{
						int id = csvProjects.GetField<int>("id");
						string title = csvProjects.GetField("title");
						string team = csvProjects.GetField("team_name");
						string url = csvProjects.GetField("url");

						url = url.Replace("http://", "https://");

						string location = csvProjects.GetField("field_location");
						string state = csvProjects.GetField("field_jurisdiction");
						string source = csvProjects.GetField("source_url");
						string video = csvProjects.GetField("video_url");

						Project project = new Project(id, title, team, url, location, state, video, source);
						projects.Add(url, project);
					}
				}
			}

			using (StreamReader srProjects = new StreamReader("Input/scraped-projects.csv"))
			{
				using (CsvReader csvProjects = new CsvReader(srProjects))
				{
					csvProjects.ReadHeader();
					while (csvProjects.Read())
					{
						string title = csvProjects.GetField("Project Name");
						string url = csvProjects.GetField("Project URI");
						string website = csvProjects.GetField("Project Website");
						string description = csvProjects.GetField("Project Description");

						if (projects.ContainsKey(url))
						{
							Project project = projects[url];
							if (project.title == title)
							{
								project.AddWebsite(website);
								project.AddDescription(description);
							}
							else
							{
								Console.WriteLine("{0} is not the same as {1}", project.title, title);
							}
						}
						else
						{
							Console.WriteLine("Error: URL ({0}) does not exist", url);
						}
					}
				}
			}

			using (StreamReader srChallenges = new StreamReader("Input/scraped-challenges.csv"))
			{
				using (CsvReader csvChallenges = new CsvReader(srChallenges))
				{
					csvChallenges.ReadHeader();
					while (csvChallenges.Read())
					{
						string title = csvChallenges.GetField("Project Name");
						string url = csvChallenges.GetField("Project URI");
						string challenge = csvChallenges.GetField("Challenge");

						if (projects.ContainsKey(url))
						{
							Project project = projects[url];
							if (project.title == title)
							{
								project.AddChallenge(challenge);
							}
							else
							{
								Console.WriteLine("{0} is not the same as {1}", project.title, title);
							}
						}
						else
						{
							Console.WriteLine("Error: URL ({0}) does not exist", url);
						}
					}
				}
			}

			using (StreamReader srJustifications = new StreamReader("Input/scraped-justifications.csv"))
			{
				using (CsvReader csvJustifications = new CsvReader(srJustifications))
				{
					csvJustifications.ReadHeader();
					while (csvJustifications.Read())
					{
						string title = csvJustifications.GetField("Project");
						string url = csvJustifications.GetField("Project URI");
						string challenge = csvJustifications.GetField("Award");
						string justification = csvJustifications.GetField("Justification");

						if (projects.ContainsKey(url))
						{
							Project project = projects[url];
							if (project.title == title)
							{
								project.AddJustification(challenge, justification);
							}
							else
							{
								Console.WriteLine("{0} is not the same as {1}", project.title, title);
							}
						}
						else
						{
							Console.WriteLine("Error: URL ({0}) does not exist", url);
						}
					}
				}
			}

			Console.WriteLine("Number of Projects: {0}", projects.Count);
        }

        public static void ReadHackerspaceVideoUhh() {
            
        }

        public static void ReadAwards() {
            awards = new Dictionary<string, Award>();

			/**
             * https://2017.hackerspace.govhack.org/awards.csv
             */

			using (StreamReader srAwards = new StreamReader("Input/hackerspace-awards.csv"))
			{
				using (CsvReader csvAwards = new CsvReader(srAwards))
				{
					csvAwards.ReadHeader();
					while (csvAwards.Read())
					{
						string name = csvAwards.GetField("name");
						string eligibility = csvAwards.GetField("eligibility_criteria");
						string sponsor = csvAwards.GetField("sponsor");
						string jurisdiction = csvAwards.GetField("jurisdiction");

						Award award = new Award(name, eligibility, sponsor, jurisdiction);
						awards.Add(name, award);
					}
				}
			}

			using (StreamReader srAwards = new StreamReader("Input/scraped-awards.csv"))
			{
				using (CsvReader csvAwards = new CsvReader(srAwards))
				{
					csvAwards.ReadHeader();
					while (csvAwards.Read())
					{
						string name = csvAwards.GetField("Award Name");
						string url = csvAwards.GetField("Award URI");
						string description = csvAwards.GetField("Award Description");

						if (awards.ContainsKey(name))
						{
							Award award = awards[name];
							award.AddURL(url);
							award.AddDescription(description);
						}
						else
						{
							Console.WriteLine("Error: URL ({0}) does not exist", url);
						}
					}
				}
			}

			foreach (KeyValuePair<string, Project> kvp in projects)
			{
				Project p = kvp.Value;
				foreach (string challenge in p.challenges)
				{
					awards[challenge].AddProject(p);
				}
			}

			Console.WriteLine("Number of Awards: {0}", awards.Count);
        }

        public static void CreateJudgingSheets() {
			foreach (KeyValuePair<string, Project> kvp in projects)
			{
				//Console.WriteLine(kvp.Value.id);
				//Console.WriteLine(kvp.Value.url);
				//continue;
				Project p = kvp.Value;
				int numAwards = 0;
				FileInfo newFile = new FileInfo("Output/" + p.safeTitle + "_" + p.id + ".xlsx");
				if (newFile.Exists)
				{
					newFile.Delete();
					newFile = new FileInfo("Output/" + p.safeTitle + "_" + p.id + ".xlsx");
				}

				using (ExcelPackage package = new ExcelPackage(newFile))
				{
					ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Judging");

					#region project start info
					//worksheet.InsertRow(1, 1);

					worksheet.Column(1).Width = 50;
					worksheet.Cells[1, 1, 4, 2].Style.Font.Size = 16;
					worksheet.Cells[1, 1, 4, 1].Style.Font.Bold = true;
					worksheet.Cells[1, 1, 4, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin);
					worksheet.Cells[1, 1].Value = "Project Title";
					worksheet.Cells[2, 1].Value = "Team Name";
					worksheet.Cells[3, 1].Value = "State, Territory or Country";
					worksheet.Cells[4, 1].Value = "Event Location";

					worksheet.Cells[1, 2, 1, 15].Merge = true;
					worksheet.Cells[2, 2, 2, 15].Merge = true;
					worksheet.Cells[3, 2, 3, 15].Merge = true;
					worksheet.Cells[4, 2, 4, 15].Merge = true;
					worksheet.Cells[1, 2, 1, 15].Value = p.title;
					worksheet.Cells[2, 2, 2, 15].Value = p.team;
					worksheet.Cells[3, 2, 3, 15].Value = p.state;
					worksheet.Cells[4, 2, 4, 15].Value = p.location;

					#endregion

					#region project description

					worksheet.Cells[6, 1, 6, 1].Style.Font.Size = 16;
					worksheet.Cells[6, 1, 6, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
					worksheet.Cells[6, 1, 6, 1].Style.Font.Bold = true;

					worksheet.Cells[6, 1, 6, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin);
					worksheet.Cells[6, 1].Value = "Description\n\n\nSome long descriptions may overflow the cell. Please check Hackerspace for full description.";

					worksheet.Cells[6, 2, 6, 15].Merge = true;
					worksheet.Cells[6, 2, 6, 15].Value = p.description;
					worksheet.Cells[6, 1, 6, 15].Style.WrapText = true;
					worksheet.Cells[6, 2, 6, 15].Style.ShrinkToFit = true;

					worksheet.Row(6).Height = 400;

					#endregion

					#region project urls

					worksheet.Cells[8, 1, 11, 2].Style.Font.Size = 16;
					worksheet.Cells[8, 1, 11, 1].Style.Font.Bold = true;

					worksheet.Cells[8, 1, 11, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin);
					worksheet.Cells[8, 1].Value = "Hackerspace URL";
					worksheet.Cells[9, 1].Value = "Project Website";
					worksheet.Cells[10, 1].Value = "Project Source";
					worksheet.Cells[11, 1].Value = "Project Video";

					worksheet.Cells[8, 2, 8, 15].Merge = true;
					worksheet.Cells[9, 2, 9, 15].Merge = true;
					worksheet.Cells[10, 2, 10, 15].Merge = true;
					worksheet.Cells[11, 2, 11, 15].Merge = true;

					try
					{
						worksheet.Cells[8, 2, 8, 15].Hyperlink = new Uri(p.url);
						worksheet.Cells[8, 2, 8, 15].Style.Font.UnderLine = true;
					}
					catch (Exception)
					{
						worksheet.Cells[8, 2, 8, 15].Value = p.url;
					}
					try
					{
						worksheet.Cells[9, 2, 9, 15].Hyperlink = new Uri(p.website);
						worksheet.Cells[9, 2, 9, 15].Style.Font.UnderLine = true;
					}
					catch (Exception)
					{
						worksheet.Cells[9, 2, 9, 15].Value = p.website;
					}
					try
					{
						worksheet.Cells[10, 2, 10, 15].Hyperlink = new Uri(p.source);
						worksheet.Cells[10, 2, 10, 15].Style.Font.UnderLine = true;
					}
					catch (Exception)
					{
						worksheet.Cells[10, 2, 10, 15].Value = p.source;
					}
					try
					{
						worksheet.Cells[11, 2, 11, 15].Hyperlink = new Uri(p.video);
						worksheet.Cells[11, 2, 11, 15].Style.Font.UnderLine = true;
					}
					catch (Exception)
					{
						worksheet.Cells[11, 2, 11, 15].Value = p.video;
					}

					#endregion

					#region general marks

					worksheet.Cells[13, 1, 13, 15].Merge = true;
					worksheet.Cells[13, 1, 13, 15].Style.Font.Size = 16;
					worksheet.Cells[13, 1, 13, 15].Style.Font.Bold = true;


					worksheet.Cells[13, 1, 13, 15].Value = "General Cross Award Judging Criteria";
					worksheet.Cells[13, 1, 13, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin);

					worksheet.Cells[14, 1, 14, 15].Style.Font.Size = 20;
					worksheet.Cells[14, 1, 19, 15].Style.Font.Size = 16;
					//worksheet.Cells[14, 1, 19, 1].Style.Font.Bold = true;
					worksheet.Cells[14, 1, 14, 15].Style.Font.Bold = true;
					worksheet.Cells[14, 1, 19, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;
					worksheet.Cells[14, 1, 19, 15].Style.Border.Left.Style = ExcelBorderStyle.Thin;
					worksheet.Cells[14, 1, 19, 15].Style.Border.Right.Style = ExcelBorderStyle.Thin;
					worksheet.Cells[14, 1, 19, 15].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

					worksheet.Cells[14, 1].Value = "Criteria";
					worksheet.Cells[14, 1, 14, 5].Merge = true;
					worksheet.Cells[14, 6].Value = "Mark (out of 10)";
					worksheet.Cells[14, 6, 14, 8].Merge = true;
					worksheet.Cells[14, 9].Value = "Comments";
					worksheet.Cells[14, 9, 14, 15].Merge = true;

					worksheet.Cells[15, 1].Value = "Originality";
					worksheet.Cells[15, 1, 15, 5].Merge = true;
					worksheet.Cells[15, 6, 15, 8].Merge = true;
					worksheet.Cells[15, 9, 15, 15].Merge = true;
					worksheet.Cells[15, 1, 15, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					worksheet.Cells[15, 6, 15, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					worksheet.Cells[16, 1].Value = "Consistency with contest purposes including social value";
					worksheet.Cells[16, 1, 16, 5].Merge = true;
					worksheet.Cells[16, 6, 16, 8].Merge = true;
					worksheet.Cells[16, 9, 16, 15].Merge = true;
					worksheet.Cells[16, 1, 16, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					worksheet.Cells[16, 6, 16, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					worksheet.Cells[17, 1].Value = "Quality and design (including standards compliance)";
					worksheet.Cells[17, 1, 17, 5].Merge = true;
					worksheet.Cells[17, 6, 17, 8].Merge = true;
					worksheet.Cells[17, 9, 17, 15].Merge = true;
					worksheet.Cells[17, 1, 17, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					worksheet.Cells[17, 6, 17, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					worksheet.Cells[18, 1].Value = "Usability (including documentation and ease of use)";
					worksheet.Cells[18, 1, 18, 5].Merge = true;
					worksheet.Cells[18, 6, 18, 8].Merge = true;
					worksheet.Cells[18, 9, 18, 15].Merge = true;
					worksheet.Cells[18, 1, 18, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					worksheet.Cells[18, 6, 18, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					worksheet.Cells[19, 1].Value = "Market / public relevance";
					worksheet.Cells[19, 1, 19, 5].Merge = true;
					worksheet.Cells[19, 6, 19, 8].Merge = true;
					worksheet.Cells[19, 9, 19, 15].Merge = true;
					worksheet.Cells[19, 1, 19, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
					worksheet.Cells[19, 6, 19, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

					worksheet.Cells[20, 6, 20, 8].Merge = true;
					worksheet.Cells[20, 6, 20, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin);
					worksheet.Cells[20, 6, 20, 8].Style.Font.Size = 20;
					worksheet.Cells[20, 6, 20, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					worksheet.Cells[20, 6, 20, 8].Formula = "=SUM(" + worksheet.Cells[15, 6].Address + ":" + worksheet.Cells[19, 6].Address + ")";

					var validation = worksheet.DataValidations.AddIntegerValidation(worksheet.Cells[15, 6].Address + ":" + worksheet.Cells[19, 6].Address);
					validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
					validation.PromptTitle = "Enter an integer value here";
					validation.Prompt = "Value should be between 0 and 10";
					validation.ShowInputMessage = true;
					validation.ErrorTitle = "An invalid value was entered";
					validation.Error = "Value must be between 0 and 10";
					validation.ShowErrorMessage = true;
					validation.Operator = ExcelDataValidationOperator.between;
					validation.Formula.Value = 0;
					validation.Formula2.Value = 10;

					worksheet.Cells[15, 6, 19, 15].Style.Locked = false;
					worksheet.Cells[15, 6, 19, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
					worksheet.Cells[15, 6, 19, 15].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);

					#endregion

					worksheet.Cells[22, 1, 22, 15].Merge = true;
					worksheet.Cells[22, 1, 22, 15].Style.Font.Size = 16;
					worksheet.Cells[22, 1, 22, 15].Style.Font.Bold = true;

					worksheet.Cells[22, 1, 22, 15].Value = "Specific Award Judging Criteria";
					worksheet.Cells[22, 1, 22, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin);

					#region awards

					int startingRow = 26;

					worksheet.InsertRow(1, 3);

					foreach (string challenge in p.challenges)
					{
						if (awards.ContainsKey(challenge))
						{
							Award award = awards[challenge];
							if (award.jurisdiction != "National/International")
							{
								continue;
							}
							numAwards++;

							worksheet.Cells[startingRow + 0, 1, startingRow + 5, 2].Style.Font.Size = 16;
							worksheet.Cells[startingRow + 2, 2, startingRow + 3, 2].Style.Font.Size = 11;
							worksheet.Cells[startingRow + 5, 2, startingRow + 5, 2].Style.Font.Size = 11;
							worksheet.Cells[startingRow + 0, 1, startingRow + 5, 1].Style.Font.Bold = true;
							worksheet.Cells[startingRow + 0, 1, startingRow + 5, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin);
							worksheet.Cells[startingRow + 0, 1].Value = "Award";
							worksheet.Cells[startingRow + 1, 1].Value = "Sponsor";
							worksheet.Cells[startingRow + 2, 1].Value = "Description\n\nSome long descriptions may overflow the cell. Please check Hackerspace for full description.";
							worksheet.Cells[startingRow + 3, 1].Value = "Eligibility Criteria\n\nSome long descriptions may overflow the cell. Please check Hackerspace for full description.";
							worksheet.Cells[startingRow + 4, 1].Value = "Hackerspace URL";
							worksheet.Cells[startingRow + 5, 1].Value = "Justification for Entry";
							worksheet.Cells[startingRow + 2, 1].Style.WrapText = true;
							worksheet.Cells[startingRow + 3, 1].Style.WrapText = true;
							worksheet.Cells[startingRow + 5, 1].Style.WrapText = true;
							worksheet.Cells[startingRow + 2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
							worksheet.Cells[startingRow + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
							worksheet.Cells[startingRow + 5, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
							worksheet.Row(startingRow + 2).Height = 100;
							worksheet.Row(startingRow + 3).Height = 100;
							worksheet.Row(startingRow + 5).Height = 100;

							worksheet.Cells[startingRow + 0, 2].Value = award.name;
							worksheet.Cells[startingRow + 1, 2].Value = award.sponsor;
							worksheet.Cells[startingRow + 2, 2].Value = award.description;
							worksheet.Cells[startingRow + 3, 2].Value = award.eligibility;

							foreach (Tuple<string, string> justification in p.justifications)
							{
								if (justification.Item1 == award.name)
								{
									if (worksheet.Cells[startingRow + 5, 2].Value == null)
									{
										worksheet.Cells[startingRow + 5, 2].Value = justification.Item2;
									}
									else
									{
										worksheet.Cells[startingRow + 5, 2].Value += "\n" + justification.Item2;
									}
								}
							}

							worksheet.Cells[startingRow + 2, 2].Style.WrapText = true;
							worksheet.Cells[startingRow + 3, 2].Style.WrapText = true;
							worksheet.Cells[startingRow + 5, 2].Style.WrapText = true;
							worksheet.Cells[startingRow + 2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
							worksheet.Cells[startingRow + 3, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
							worksheet.Cells[startingRow + 5, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
							worksheet.Cells[startingRow + 4, 2].Hyperlink = new Uri(award.url);
							worksheet.Cells[startingRow + 4, 2].Style.Font.UnderLine = true;

							worksheet.Cells[startingRow + 0, 2, startingRow + 0, 15].Merge = true;
							worksheet.Cells[startingRow + 1, 2, startingRow + 1, 15].Merge = true;
							worksheet.Cells[startingRow + 2, 2, startingRow + 2, 15].Merge = true;
							worksheet.Cells[startingRow + 3, 2, startingRow + 3, 15].Merge = true;
							worksheet.Cells[startingRow + 4, 2, startingRow + 4, 15].Merge = true;
							worksheet.Cells[startingRow + 5, 2, startingRow + 5, 15].Merge = true;


							worksheet.Cells[startingRow + 6, 1, startingRow + 6, 15].Style.Font.Size = 20;
							worksheet.Cells[startingRow + 6, 1, startingRow + 9, 15].Style.Font.Size = 16;
							//worksheet.Cells[startingRow + 5, 1, startingRow + 9, 1].Style.Font.Bold = true;
							worksheet.Cells[startingRow + 6, 1, startingRow + 6, 15].Style.Font.Bold = true;
							worksheet.Cells[startingRow + 6, 1, startingRow + 9, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;
							worksheet.Cells[startingRow + 6, 1, startingRow + 9, 15].Style.Border.Left.Style = ExcelBorderStyle.Thin;
							worksheet.Cells[startingRow + 6, 1, startingRow + 9, 15].Style.Border.Right.Style = ExcelBorderStyle.Thin;
							worksheet.Cells[startingRow + 6, 1, startingRow + 9, 15].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

							worksheet.Cells[startingRow + 6, 1].Value = "Criteria";
							worksheet.Cells[startingRow + 6, 1, startingRow + 6, 5].Merge = true;
							worksheet.Cells[startingRow + 6, 6].Value = "Mark (out of 10)";
							worksheet.Cells[startingRow + 6, 6, startingRow + 6, 8].Merge = true;
							worksheet.Cells[startingRow + 6, 9].Value = "Comments";
							worksheet.Cells[startingRow + 6, 9, startingRow + 6, 15].Merge = true;

							worksheet.Cells[startingRow + 7, 1].Value = "General Criteria (Carried Down)";
							worksheet.Cells[startingRow + 8, 1].Value = "The relevance to the team nominated category definition";
							worksheet.Cells[startingRow + 9, 1].Value = "Specific prize eligibility criteria detailed (if any) e.g. data use, team criteria";

							worksheet.Cells[startingRow + 7, 1, startingRow + 9, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

							worksheet.Cells[startingRow + 7, 1, startingRow + 7, 5].Merge = true;
							worksheet.Cells[startingRow + 8, 1, startingRow + 8, 5].Merge = true;
							worksheet.Cells[startingRow + 9, 1, startingRow + 9, 5].Merge = true;
							worksheet.Cells[startingRow + 7, 6, startingRow + 7, 8].Merge = true;
							worksheet.Cells[startingRow + 8, 6, startingRow + 8, 8].Merge = true;
							worksheet.Cells[startingRow + 9, 6, startingRow + 9, 8].Merge = true;
							worksheet.Cells[startingRow + 7, 9, startingRow + 7, 15].Merge = true;
							worksheet.Cells[startingRow + 8, 9, startingRow + 8, 15].Merge = true;
							worksheet.Cells[startingRow + 9, 9, startingRow + 9, 15].Merge = true;

							worksheet.Cells[startingRow + 10, 6, startingRow + 10, 8].Merge = true;
							worksheet.Cells[startingRow + 10, 6, startingRow + 10, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin);
							worksheet.Cells[startingRow + 10, 6, startingRow + 10, 8].Style.Font.Size = 16;
							worksheet.Cells[startingRow + 10, 6, startingRow + 10, 8].Style.Font.Bold = true;

							worksheet.Cells[startingRow + 7, 6, startingRow + 7, 8].Formula = "=" + worksheet.Cells[23, 6].Address;
							worksheet.Cells[startingRow + 10, 6, startingRow + 10, 8].Formula = "=SUM(" + worksheet.Cells[startingRow + 7, 6].Address + ":" + worksheet.Cells[startingRow + 9, 6].Address + ")";

							worksheet.Cells[startingRow + 7, 6, startingRow + 10, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

							worksheet.Cells[startingRow + 8, 6, startingRow + 9, 15].Style.Locked = false;
							worksheet.Cells[startingRow + 8, 6, startingRow + 9, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
							worksheet.Cells[startingRow + 8, 6, startingRow + 9, 15].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);

							validation = worksheet.DataValidations.AddIntegerValidation(worksheet.Cells[startingRow + 8, 6].Address + ":" + worksheet.Cells[startingRow + 9, 6].Address);
							validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
							validation.PromptTitle = "Enter an integer value here";
							validation.Prompt = "Value should be between 0 and 10";
							validation.ShowInputMessage = true;
							validation.ErrorTitle = "An invalid value was entered";
							validation.Error = "Value must be between 0 and 10";
							validation.ShowErrorMessage = true;
							validation.Operator = ExcelDataValidationOperator.between;
							validation.Formula.Value = 0;
							validation.Formula2.Value = 10;

							startingRow += 12;
						}
						else
						{
							Console.WriteLine(challenge + " is lost");
						}
						//break;
					}

					#endregion

					//worksheet.InsertRow(1, 1);

					worksheet.Cells[1, 1, 2, 2].Style.Font.Size = 16;
					worksheet.Cells[1, 1, 2, 1].Style.Font.Bold = true;
					worksheet.Cells[1, 1, 2, 15].Style.Border.Top.Style = ExcelBorderStyle.Thin;
					worksheet.Cells[1, 1, 2, 15].Style.Border.Left.Style = ExcelBorderStyle.Thin;
					worksheet.Cells[1, 1, 2, 15].Style.Border.Right.Style = ExcelBorderStyle.Thin;
					worksheet.Cells[1, 1, 2, 15].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
					worksheet.Cells[1, 1].Value = "Judge's Name";
					worksheet.Cells[2, 1].Value = "Judge's Email";
					worksheet.Cells[1, 2, 2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
					worksheet.Cells[1, 2, 2, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
					worksheet.Cells[1, 2, 2, 2].Style.Locked = false;

					worksheet.Cells[1, 2, 1, 15].Merge = true;
					worksheet.Cells[2, 2, 2, 15].Merge = true;

					worksheet.Protection.AllowInsertRows = false;
					worksheet.Protection.AllowSort = false;
					worksheet.Protection.AllowSelectLockedCells = true;
					worksheet.Protection.AllowSelectUnlockedCells = true;
					worksheet.Protection.AllowAutoFilter = false;
					worksheet.Protection.AllowInsertColumns = false;
					worksheet.Protection.IsProtected = true;

					package.Save();
				}

				if (numAwards == 0)
				{
					//Console.WriteLine(p.title);
					newFile.Delete();

				}
				else
				{
					if (p.video == "")
					{
						//Console.WriteLine(p.url);
					}
					newFile.MoveTo("Output/" + p.safeTitle + "___id-" + p.id + "___n-" + numAwards + ".xlsx");
					p.AddFilename(p.safeTitle + "___id-" + p.id + "___n-" + numAwards + ".xlsx");
				}

				//break;
			}
        }
    
        public static void ReadJudgedSheets() {
			string[] judgedFiles = Directory.GetFiles("InputMarked");

			foreach (string judgedFile in judgedFiles)
			{
				FileInfo jFileInfo = new FileInfo(judgedFile);

				using (ExcelPackage package = new ExcelPackage(jFileInfo))
				{
					ExcelWorksheet worksheet = package.Workbook.Worksheets["Judging"];

					string jName = (string)worksheet.Cells[1, 2].Value;
					string jEmail = (string)worksheet.Cells[2, 2].Value;

					string jProject = (string)worksheet.Cells[4, 2].Value;
					string jURL = (string)worksheet.Cells[11, 2].Value;

					projects[jURL].AddJudging(jName, jEmail);
					//Console.WriteLine(jURL);
					if (worksheet.Cells[18, 6].Value == null || worksheet.Cells[19, 6].Value == null ||
						worksheet.Cells[20, 6].Value == null || worksheet.Cells[21, 6].Value == null ||
						worksheet.Cells[22, 6].Value == null)
					{
						//Console.WriteLine(jProject);
						//Console.WriteLine(judgedFile);
						//Console.WriteLine("-------");
						continue;
					}

					float originality = Convert.ToSingle(worksheet.Cells[18, 6].Value);
					float consistency = Convert.ToSingle(worksheet.Cells[19, 6].Value);
					float quality = Convert.ToSingle(worksheet.Cells[20, 6].Value);
					float usability = Convert.ToSingle(worksheet.Cells[21, 6].Value);
					float relevance = Convert.ToSingle(worksheet.Cells[22, 6].Value);

					projects[jURL].judging.AddOriginality(originality);
					projects[jURL].judging.AddConsistency(consistency);
					projects[jURL].judging.AddQuality(quality);
					projects[jURL].judging.AddUsability(usability);
					projects[jURL].judging.AddRelevance(relevance);

					//Console.WriteLine(projects[jURL].filename);

					int row = 26;

					while (true)
					{
						if ((string)worksheet.Cells[row, 1].Value == "Award")
						{

							if (worksheet.Cells[row + 8, 6].Value == null || worksheet.Cells[row + 9, 6].Value == null)
							{
								//Console.WriteLine(jProject);
								//Console.WriteLine(judgedFile);
								//Console.WriteLine("-------");
								break;
							}

							string jAward = (string)worksheet.Cells[row, 2].Value;
							//Console.WriteLine(jAward);
							float teamRelevance = Convert.ToSingle(worksheet.Cells[row + 8, 6].Value);
							float specific = Convert.ToSingle(worksheet.Cells[row + 9, 6].Value);
							projects[jURL].judging.AddAward(jAward, teamRelevance, specific);
							row += 12;
						}
						else
						{
							break;
						}
					}

					if (projects[jURL].challenges.Count > 7)
					{
						foreach (KeyValuePair<string, Tuple<float, float>> pair in projects[jURL].judging.awardJudging)
						{
							if ((pair.Value.Item1 + pair.Value.Item2) < 7)
							{
								projects[jURL].judging.Penalty();
							}
						}
					}
				}
			}
        }
    
        public static void OutputJudgedResults() {

            using (StreamWriter srJudged = new StreamWriter("JudgedEntries.csv")) {
                using (CsvWriter csvJudged = new CsvWriter(srJudged)) {
                    csvJudged.WriteField("Award");
                    csvJudged.WriteField("Project");
                    csvJudged.WriteField("Project URI");
                    csvJudged.WriteField("General Result");
                    csvJudged.WriteField("Award Result");
                    csvJudged.WriteField("Total Result");
                    csvJudged.WriteField("Over Selection Penalty");
                    csvJudged.WriteField("Over Selection Result");
                    csvJudged.WriteField("General - Originality");
                    csvJudged.WriteField("General - Consistency");
                    csvJudged.WriteField("General - Relevance");
                    csvJudged.WriteField("General - Quality");
                    csvJudged.WriteField("General - Usability");
                    csvJudged.WriteField("Award - Relevance");
                    csvJudged.WriteField("Award - Specific");
                    csvJudged.NextRecord();

                    foreach (KeyValuePair<string, Award> kvp in awards) {
                        Award award = kvp.Value;

                        if (award.jurisdiction != "National/International") {
                            continue;
                        }

                        foreach (Project p in award.projects) {
                            if (p.judging != null && p.judging.awardJudging.ContainsKey(award.name)) {
                                csvJudged.WriteField(award.name);
                                csvJudged.WriteField(p.title);
                                csvJudged.WriteField(p.url);

                                float resultGeneral = 0;
                                resultGeneral += p.judging.originality;
                                resultGeneral += p.judging.consistency;
                                resultGeneral += p.judging.relevance;
                                resultGeneral += p.judging.quality;
                                resultGeneral += p.judging.usability;

                                float resultAward = 0;
                                resultAward += p.judging.awardJudging[award.name].Item1;
                                resultAward += p.judging.awardJudging[award.name].Item2;

                                csvJudged.WriteField(resultGeneral);
                                csvJudged.WriteField(resultAward);
                                csvJudged.WriteField(resultGeneral + resultAward);
                                csvJudged.WriteField(p.judging.overSelection);
                                csvJudged.WriteField(resultGeneral + resultAward + p.judging.overSelection);
                                csvJudged.WriteField(p.judging.originality);
                                csvJudged.WriteField(p.judging.consistency);
                                csvJudged.WriteField(p.judging.relevance);
                                csvJudged.WriteField(p.judging.quality);
                                csvJudged.WriteField(p.judging.usability);
								csvJudged.WriteField(p.judging.awardJudging[award.name].Item1);
								csvJudged.WriteField(p.judging.awardJudging[award.name].Item2);
                            } else {
								csvJudged.WriteField(award.name);
								csvJudged.WriteField(p.title);
								csvJudged.WriteField(p.url);

                                csvJudged.WriteField("-");
                                csvJudged.WriteField("-");
								csvJudged.WriteField("-");
								csvJudged.WriteField("-");
                                csvJudged.WriteField("-");
                                csvJudged.WriteField("-");
                                csvJudged.WriteField("-");
                                csvJudged.WriteField("-");
                                csvJudged.WriteField("-");
                                csvJudged.WriteField("-");
                                csvJudged.WriteField("-");
                                csvJudged.WriteField("-");
                            }
                            csvJudged.NextRecord();
                        }
                    }
                }
            }
        }
    
        public static void BaselineJudges() {
			using (StreamReader srJudges = new StreamReader("JudgeTracking.csv"))
			{
				using (CsvReader csvJudges = new CsvReader(srJudges))
				{
					csvJudges.ReadHeader();
					while (csvJudges.Read())
					{
                        string projectFile = csvJudges.GetField("Project");
                        string judgeName = csvJudges.GetField("Judge A");

                        if (judgeName != "")
                        {
                            foreach (KeyValuePair<string, Project> kvp in projects) {
                                if (projectFile == kvp.Value.filename) {
                                    kvp.Value.assignedJudge = judgeName;
                                    break;
                                }
                            }
                        }
					}
				}
			}

            Dictionary<string, List<float>> judgeResults = new Dictionary<string, List<float>>();

            foreach (KeyValuePair<string, Project> kvp in projects) {
                if (kvp.Value.assignedJudge != null && kvp.Value.assignedJudge != "") {
                    if (judgeResults.ContainsKey(kvp.Value.assignedJudge) == false) {
                        judgeResults.Add(kvp.Value.assignedJudge, new List<float>());
                    }
                }

                if (kvp.Value.judging == null) {
                    continue;
                }

                float results = 0;
                results += kvp.Value.judging.originality;
                results += kvp.Value.judging.consistency;
                results += kvp.Value.judging.relevance;
                results += kvp.Value.judging.quality;
                results += kvp.Value.judging.usability;
                judgeResults[kvp.Value.assignedJudge].Add(results);
            }

            Dictionary<string, float> judgeBaseline = new Dictionary<string, float>();

            foreach (KeyValuePair<string, List<float>> kvp in judgeResults) {
                float results = 0;
                foreach (float f in kvp.Value) {
                    results += f;
                }
                results = results / kvp.Value.Count;
                judgeBaseline.Add(kvp.Key, results);
            }

            using (StreamWriter swBaseline = new StreamWriter("Baseline.csv")) {
                using (CsvWriter csvBaseline = new CsvWriter(swBaseline)) {
                    csvBaseline.WriteField("Judge");
                    csvBaseline.WriteField("Scored Average");
                    csvBaseline.WriteField("Number Judged");
                    csvBaseline.NextRecord();

                    foreach (KeyValuePair<string, float> kvp in judgeBaseline) {
                        csvBaseline.WriteField(kvp.Key);
                        csvBaseline.WriteField(kvp.Value);
                        csvBaseline.WriteField(judgeResults[kvp.Key].Count);
                        csvBaseline.NextRecord();
                    }
                }
            }
        }
    }
}
