using System.Collections.Generic;

namespace SlingMD.Outlook.Services.Formatting
{
    internal class MatchResult
    {
        public MatchTier Tier { get; }
        public IReadOnlyList<ContactIndexEntry> Candidates { get; }

        public MatchResult(MatchTier tier, IReadOnlyList<ContactIndexEntry> candidates)
        {
            Tier = tier;
            Candidates = candidates;
        }
    }
}
