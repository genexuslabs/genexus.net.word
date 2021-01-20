using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Genexus.Word
{

    public interface ISearchTextReplacer
    {
        int ReplaceText(Paragraph paragraph, string find, string replaceWith, bool matchCase, List<string> properties);
    }


    public class SearchTextReplaceStrategy: ISearchTextReplacer
    {
        /// <summary>
        /// Find/replace within the specified paragraph.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="find"></param>
        /// <param name="replaceWith"></param>
        public int ReplaceText(Paragraph paragraph, string find, string replaceWith, bool matchCase, List<string> properties)
        {
            Dictionary<Run, List<Run>> addedRuns = new Dictionary<Run, List<Run>>();

            int replaceCount = 0;
            var texts = paragraph.Descendants<Text>();
            for (int t = 0; t < texts.Count(); t++)
            {   // figure out which Text element within the paragraph contains the starting point of the search string
                Text startTxt = texts.ElementAt(t);
                for (int c = 0; c < startTxt.Text.Length; c++)
                {
                    var match = IsMatch(texts, t, c, find, matchCase);
                    if (match != null)
                    {   // now replace the text
                        string[] lines = replaceWith.Replace(Environment.NewLine, "\r").Split('\n', '\r'); // handle any lone n/r returns, plus newline.

                        int skip = lines[lines.Length - 1].Length - 1; // will jump to end of the replacement text, it has been processed.

                        if (c > 0)
                        {
                            lines[0] = startTxt.Text.Substring(0, c) + lines[0];  // has a prefix
                        }
                        if (match.EndCharIndex + 1 < texts.ElementAt(match.EndElementIndex).Text.Length)
                        {
                            lines[lines.Length - 1] = lines[lines.Length - 1] + texts.ElementAt(match.EndElementIndex).Text.Substring(match.EndCharIndex + 1);
                        }

                        int replaceIdx = c;

                        startTxt.Text = lines[0].Substring(0, replaceIdx);
                        startTxt.Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve); // in case your value starts/ends with whitespace;
                        string remainingText = lines[0].Substring(replaceIdx);

                        replaceCount += 1;
                        // remove any extra texts.
                        for (int i = t + 1; i <= match.EndElementIndex; i++)
                        {
                            texts.ElementAt(i).Text = string.Empty; // clear the text
                        }

                        RunProperties runProps;
                        Run parentRun;
                        TryGetRunProperties(startTxt, out runProps, out parentRun);

                        CreateAppendText(remainingText.Substring(replaceWith.Length), runProps, parentRun, addedRuns);
                        CreateAppendText(replaceWith, new RunProperties(Helper.GetProperties(properties)), parentRun, addedRuns);

                        // if 'with' contained line breaks we need to add breaks back...
                        if (lines.Count() > 1)
                        {
                            OpenXmlElement currEl = startTxt;
                            Break br;

                            // append more lines
                            var run = startTxt.Parent as Run;
                            for (int i = 1; i < lines.Count(); i++)
                            {
                                br = new Break();
                                run.InsertAfter<Break>(br, currEl);
                                currEl = br;
                                startTxt = new Text(lines[i]);
                                run.InsertAfter<Text>(startTxt, currEl);
                                t++; // skip to this next text element
                                currEl = startTxt;
                            }
                            c = skip; // new line
                        }
                        else
                        {   // continue to process same line
                            c += skip;
                        }
                    }
                }
            }

            foreach (var item in addedRuns)
            {
                foreach (Run run in item.Value)
                {
                    paragraph.InsertAfter(run, item.Key);
                }
            }
            return replaceCount;
        }

        private static void TryGetRunProperties(Text startTxt, out RunProperties runProps, out Run parentRun)
        {
            runProps = null;
            parentRun = null;
            try
            {
                Run runContainer = (Run)startTxt.Parent;
                runProps = runContainer.GetFirstChild<RunProperties>();
                if (runProps != null)
                {
                    runProps = (RunProperties)runProps.Clone();
                }
                parentRun = (Run)startTxt.Parent;
            }
            catch (Exception)
            {

            }
        }

        private static Run CreateAppendText(string text, RunProperties runProps, Run parentRun, Dictionary<Run, List<Run>> runList)
        {
            Run r = new Run(runProps);
            Text newText = new Text(text);
            newText.Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve); // in case your value starts/ends with whitespace;
            r.Append(newText);
            if (!runList.ContainsKey(parentRun))
            {
                runList[parentRun] = new List<Run>();
            }
            runList[parentRun].Add(r);
            return r;
        }


        /// <summary>
        /// Determine if the texts (starting at element t, char c) exactly contain the find text
        /// </summary>
        /// <param name="texts"></param>
        /// <param name="t"></param>
        /// <param name="c"></param>
        /// <param name="find"></param>
        /// <returns>null or the result info</returns>
        static Match IsMatch(IEnumerable<Text> texts, int t, int c, string find, bool matchCase)
        {
            find = (matchCase) ? find : find.ToLower();
            int ix = 0;
            for (int i = t; i < texts.Count(); i++)
            {
                for (int j = c; j < texts.ElementAt(i).Text.Length; j++)
                {
                    string currentText = (matchCase) ? texts.ElementAt(i).Text : texts.ElementAt(i).Text.ToLower(); 
                    if (find[ix] != currentText[j])
                    {
                        return null; // element mismatch
                    }

                    ix++; // match; go to next character
                    if (ix == find.Length)
                    {
                        return new Match() { 
                            EndElementIndex = i, 
                            EndCharIndex = j 
                        }; // full match with no issues
                    }
                }
                c = 0; // reset char index for next text element
            }
            return null; // ran out of text, not a string match
        }

        /// <summary>
        /// Defines a match result
        /// </summary>
        class Match
        {
            /// <summary>
            /// Last matching element index containing part of the search text
            /// </summary>
            public int EndElementIndex { get; set; }
            /// <summary>
            /// Last matching char index of the search text in last matching element
            /// </summary>
            public int EndCharIndex { get; set; }
        }

    }   // class


    public class PowerToolsSearchTextReplaceStrategy : ISearchTextReplacer
    {
        /// <summary>
        /// Find/replace within the specified paragraph.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="find"></param>
        /// <param name="replaceWith"></param>
        public int ReplaceText(Paragraph paragraph, string find, string replaceWith, bool matchCase, List<string> properties)
        {
            return 0;
        }
    }
}
