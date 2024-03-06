using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects;
using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using System.Text.RegularExpressions;

namespace Appeon.DotnetDemo.Dw2Doc.Dw
{
    public class DwTools
    {
        private static Regex GetDwObjectPropertyRegex(string obj, string property) => new Regex(obj + @"\(" + property + @"=(\d+).*\)");

        private static Regex MainBandRegex(string bandName) => GetDwObjectPropertyRegex(bandName, "height");
        private static Regex GroupBandRegex() => new Regex(@"group *\(.*level=(\d+).*header\.height=(\d+).*trailer\.height=(\d+).*\)");


        public static IList<DwBand> GetBands(string dwSyntax, double yrate, string[] excludedBands)
        {

            /// will need to insert groups between summary and detail
            string[] expectedBands = { "header", "detail", "summary", "footer" };
            const int trailerBandsStart = 1;

            Match[] expectedMatches = new Match[expectedBands.Length];

            { // create scope because I want to use i again later
                int i = 0;
                foreach (var band in expectedBands)
                {
                    expectedMatches[i] = MainBandRegex(band).Match(dwSyntax);
                    if (!expectedMatches[i].Success || expectedMatches[i].Groups.Count < 2)
                    {
                        throw new ArgumentException($"Invalid syntax. Expected band: [{band}]", nameof(dwSyntax));
                    }
                    ++i;
                }
            }

            var groupBandHeaders = new List<DwBand>();
            var groupBandTrailers = new List<DwBand>();

            var groupBandMatches = GroupBandRegex().Matches(dwSyntax);
            foreach (var bandMatch in groupBandMatches.Cast<Match>())
            {
                var newHeader = new DwBand($"header.{bandMatch.Groups[1].Value}", true)
                {
                    Height = (int)(int.Parse(bandMatch.Groups[2].Value) / yrate)
                };

                var newTrailer = new DwBand($"trailer.{bandMatch.Groups[1].Value}", false)
                {
                    Height = (int)(int.Parse(bandMatch.Groups[3].Value) / yrate),

                    RelatedHeader = newHeader
                };
                newHeader.RelatedTrailers.Add(newTrailer);
                if (groupBandHeaders.Count > 0)
                {
                    newHeader.ParentBand = groupBandHeaders[groupBandHeaders.Count - 1];
                }
                groupBandHeaders.Add(newHeader);
                groupBandTrailers.Add(newTrailer);
            }

            var bands = new List<DwBand>();

            int accumulatedHeight = 0;

            int groupHeaderTailIndex = 0;
            for (int i = 0; i < expectedBands.Length; i++)
            {
                if (excludedBands.Contains(expectedBands[i]))
                    continue;
                if (i == trailerBandsStart)
                {
                    foreach (var band in groupBandHeaders)
                    { // Calculate the heights of the group bands
                        // band.Position = accumulatedHeight;
                        band.BandType = Common.Enums.BandType.Header;
                        bands.Add(band);
                        //accumulatedHeight += band.Height;
                    }

                    // save position of last group header to insert detail
                    groupHeaderTailIndex = bands.Count;

                    groupBandTrailers.Reverse();
                    foreach (var band in groupBandTrailers)
                    {
                        //band.Position = accumulatedHeight;
                        band.BandType = Common.Enums.BandType.Trailer;
                        bands.Add(band);
                        //accumulatedHeight += band.Height;
                    }
                }

                var newBand = new DwBand(expectedBands[i], expectedBands[i] == "detail")
                {
                    Height = (int)(int.Parse(expectedMatches[i].Groups[1].Value) / yrate),
                    //Position = accumulatedHeight,
                    BandType = i <= trailerBandsStart ? Common.Enums.BandType.Header : Common.Enums.BandType.Trailer
                };
                if (i > trailerBandsStart)
                {
                    bands[0].RelatedTrailers.Add(newBand);
                    newBand.RelatedHeader = bands[0];
                }
                if (expectedBands[i] == "detail")
                {
                    bands.Insert(groupHeaderTailIndex, newBand);
                }
                else
                    bands.Add(newBand);

                // accumulatedHeight += newBand.Height;
            }

            // Calculate band offsets
            foreach (var band in bands)
            {
                band.Position = accumulatedHeight;
                accumulatedHeight += band.Height;
            }

            return bands;
        }

        public static int GetDwType(string typeIndex, out DwType type, out string? error)
        {
            type = default;
            error = null;

            if (int.TryParse(typeIndex, out int parsed))
            {
                type = (DwType)parsed;

            }
            else
            {
                error = "Could not parse input value";
                return -1;
            }

            return 0;
        }

        public static Alignment GetAlignment(short alignment) => (Alignment)alignment;
    }
}
