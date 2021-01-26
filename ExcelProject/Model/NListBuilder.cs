using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProject.Model
{
    class NListBuilder
    {
        public Dictionary<int, List<string>> tags { get; set; }
        public string character { get; set; }
        public NListBuilder()
        {
        }
        public NListBuilder(Dictionary<int, List<string>> tags, string character)
        {
            this.tags = tags;
            this.character = character;
        }

        public List<string> AllCombos
        {
            get
            {
                return GetCombos(tags);
            }
        }

        List<string> GetCombos(IEnumerable<KeyValuePair<int, List<string>>> remainingTags)
        {
            if (remainingTags.Count() == 1)
            {
                return remainingTags.First().Value;
            }
            else
            {
                var current = remainingTags.First();
                List<string> outputs = new List<string>();
                List<string> combos = GetCombos(remainingTags.Where(tag => tag.Key != current.Key));

                foreach (var tagPart in current.Value)
                {
                    foreach (var combo in combos)
                    {
                        outputs.Add(string.Format("{0} {1} {2}", tagPart, character, combo));
                    }
                }

                return outputs;
            }


        }
    }
}
