using System;
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

        public Award(string name, string eligibility, string sponsor, string jurisdiction)
        {
            this.name = name;
            this.eligibility = eligibility;
            this.sponsor = sponsor;
            this.jurisdiction = jurisdiction;
        }

        public void AddDescription(string description)
        {
            this.description = description;
        }

        public void AddURL(string url)
        {
            this.url = url;
        }
    }
}
