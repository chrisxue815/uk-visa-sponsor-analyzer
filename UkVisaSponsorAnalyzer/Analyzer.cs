using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.IO;
using System.Linq;

namespace UkVisaSponsorAnalyzer
{
    public class Analyzer
    {
        private StreamWriter writer;
        private bool manualMode;
        private int totalNumSponsors;

        public void Run()
        {
            var settingsStr = File.ReadAllText("settings.toml");
            var settings = Toml.Toml.Parse(settingsStr);

            var reader = new PdfReader(settings.input);
            writer = new StreamWriter(settings.output);
            manualMode = settings.manualMode;
            totalNumSponsors = 0;

            // first page
            {
                // Parameters: distanceInPixelsFromLeft, distanceInPixelsFromBottom, width, height
                var rect = new System.util.RectangleJ(24, 34, 326, 348);
                var strategy = CreateStrategy(rect);
                var text = PdfTextExtractor.GetTextFromPage(reader, 1, strategy);
                Write(text);
            }

            // middle page
            {
                var rect = new System.util.RectangleJ(24, 0, 326, 569);

                for (var i = 2; i < reader.NumberOfPages; i++)
                {
                    var strategy = CreateStrategy(rect);
                    var text = PdfTextExtractor.GetTextFromPage(reader, i, strategy);
                    text = text.Replace("Organisation Name\n", "");
                    Write(text);
                }
            }

            // last page
            {
                var rect = new System.util.RectangleJ(24, 229, 326, 339);
                var strategy = CreateStrategy(rect);
                var text = PdfTextExtractor.GetTextFromPage(reader, reader.NumberOfPages, strategy);
                Write(text);
            }

            writer.Flush();

            if (totalNumSponsors != settings.totalNumSponsors)
            {
                Console.WriteLine(
$@"Warning: mismatched total number of sponsors:
 expected `{settings.totalNumSponsors}`,
    found `{totalNumSponsors}`");
                Console.ReadLine();
            }
        }

        private static ITextExtractionStrategy CreateStrategy(System.util.RectangleJ rect)
        {
            var filters = new RenderFilter[1];
            filters[0] = new RegionTextRenderFilter(rect);

            return new FilteredTextRenderListener(
                new LocationTextExtractionStrategy(),
                filters);
        }

        private void Write(string text)
        {
            if (manualMode)
            {
                Console.WriteLine(text);
                Console.ReadLine();
            }
            writer.WriteLine(text);
            totalNumSponsors += text.Count(ch => ch == '\n') + 1;
        }
    }
}
