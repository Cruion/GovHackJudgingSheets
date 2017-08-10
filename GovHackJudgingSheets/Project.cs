using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace GovHackJudgingSheets
{
    public class Project
    {
        public int id;
        public string title;
        public string safeTitle;
        public string team;
        public string url;
        public string location;
        public string state;
        public string description; 
        public string website;
        public string source;
        public string video; 

        public List<string> challenges;
        public List<Tuple<string, string>> justifications;

        public Project(int id, string title, string team, string url, string location, string state, string video, string source)
        {
            this.id = id;
            this.title = title;

            Regex rgx = new Regex("[^a-zA-Z0-9 -]");
            this.safeTitle = rgx.Replace(title, "");

            this.team = team;
            this.url = url;
            this.location = location;
            this.state = state;
            this.source = source;
            this.video = video;

            this.challenges = new List<string>();
            this.justifications = new List<Tuple<string, string>>();
        }

        public void AddDescription(string description)
        {
            this.description = description;
        }

        public void AddWebsite(string website)
        {
            this.website = website;
        }

        public void AddChallenge(string challenge) 
        {
            this.challenges.Add(challenge);
        }

        public void AddJustification(string challenge, string justification)
        {
            this.justifications.Add(new Tuple<string, string>(challenge, justification));
        }
    }
}
