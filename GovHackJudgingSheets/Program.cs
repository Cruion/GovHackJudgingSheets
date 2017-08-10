using System;
using System.Collections.Generic;
using System.IO;
using CsvHelper;

namespace GovHackJudgingSheets
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            Dictionary<string, Project> projects = new Dictionary<string, Project>();
            Dictionary<string, Award> awards = new Dictionary<string, Award>();

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

                        if (projects.ContainsKey(url)) {
                            Project project = projects[url];
                            if (project.title == title) {
                                project.AddWebsite(website);
                                project.AddDescription(description);
                            } else {
                                Console.WriteLine("{0} is not the same as {1}", project.title, title);
                            }
                        } else {
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

            Console.WriteLine("Number of Awards: {0}", awards.Count);
        }
    }
}
