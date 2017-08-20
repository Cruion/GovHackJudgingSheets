using System;
using System.Collections.Generic;
namespace GovHackJudgingSheets
{
    public class Award
    {
        public string name;
        public string url;
        public string description;
        public string eligibility;
        public string sponsor;
        public string jurisdiction;

        public List<Project> projects;

        public Award(string name, string eligibility, string sponsor, string jurisdiction)
        {
            this.name = name;
            this.eligibility = eligibility;
            this.sponsor = sponsor;
            this.jurisdiction = jurisdiction;
            this.projects = new List<Project>();
        }

        public void AddDescription(string description)
        {
            this.description = description;
        }

        public void AddURL(string url)
        {
            this.url = url;
        }

        public void AddProject(Project project)
        {
            this.projects.Add(project);
        }
    }
}
