using System;
using System.Collections.Generic;

namespace GovHackJudgingSheets
{
    public class Judging
    {
        public string name;
        public string email;

        public float originality;
        public float consistency;
        public float quality;
        public float usability;
        public float relevance;
        public float overSelection;

        public Dictionary<string, Tuple<float, float>> awardJudging;

        public Judging(string name, string email)
        {
            this.name = name;
            this.email = email;

            this.awardJudging = new Dictionary<string, Tuple<float, float>>();
            this.overSelection = 0;
        }

        public void AddOriginality(float originality)
        {
            this.originality = originality;
        }

		public void AddConsistency(float consistency)
		{
			this.consistency = consistency;
		}

		public void AddQuality(float quality)
		{
			this.quality = quality;
		}

		public void AddUsability(float usability)
		{
			this.usability = usability;
		}

		public void AddRelevance(float relevance)
		{
			this.relevance = relevance;
		}

        public void AddAward(string award, float relevance, float specific)
        {
            this.awardJudging.Add(award, new Tuple<float, float>(relevance, specific));
        }

        public void Penalty()
        {
            if (this.overSelection == 0) {
                this.overSelection = -1;
            } else {
                this.overSelection = this.overSelection + (this.overSelection * 0.25f);
            }
        }
    }
}
