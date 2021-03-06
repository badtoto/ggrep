﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Text.RegularExpressions;
using System.ComponentModel;
using GGrep.Instance;

namespace GGrep.Pattern
{
    abstract class BasePattern
    {
        protected GForm parent = null;

        int MAX_WIDTH = 127;

        public BasePattern() { }

        public BasePattern(GForm _parent)
        {
            parent = _parent;
        }

        protected Regex GetRegex()
        {
            RegexOptions ro = RegexOptions.Singleline;
            string input = parent.Option.SearchString;

            if (parent.Option.IsRegex)
            {
                ro = ro | RegexOptions.IgnoreCase;
#if DEBUG1
                        if (option.Multiline)
                        {
                            ro = ro | RegexOptions.Multiline;
                        }
#endif
            }
            else
            {
                if (!parent.Option.IsCaseSensitive)
                {
                    ro = ro | RegexOptions.IgnoreCase;
                }

                input = Utils.SearchStringRegexEscaped(parent.Option.SearchString);
                if (parent.Option.IsSearchOnWords)
                {
                    input = @"\b" + input + @"\b";
                }
            }
            
            return new Regex(input, ro);
        }

        /// <summary>
        /// Search in line
        /// </summary>
        /// <param name="line">txt</param>
        /// <param name="list">result data</param>
        /// <param name="path">file path</param>
        /// <param name="rowNo">row No.</param>
        /// <param name="encoding">encoding</param>
        /// <param name="regex">regex</param>
        public void AnalyzeLine(string line, ArrayList list, string path, long rowNo, string encoding, Regex regex)
        {
            // for fast search
            if (!parent.Option.IsRegex)
            {
                if (parent.Option.IsCaseSensitive)
                {
                    if (!line.Contains(parent.Option.SearchString))
                        return;
                }
                else
                {
                    if (!line.ToLower().Contains(parent.Option.SearchString.ToLower()))
                        return;
                }
            }

            foreach (Match m in regex.Matches(line))
            {
                ResultData data = new ResultData();
                data.SelectedPath = parent.Option.SearchFolder;
                data.No = (++parent.Status.Hit);
                data.FullFileName = path;
                data.RowNo = rowNo;
                data.ColNo = m.Index + 1;
                data.Line = line;
                data.MatchedString = m.Value;
                data.FileEncoding = encoding;

                if (line.Length > MAX_WIDTH)
                {
                    if (data.MatchedString.Length >= line.Length)
                    {
                        data.ShowLine = data.MatchedString;
                        data.ShowColNo = 0;
                    }
                    else
                    {
                        data.ShowColNo = (MAX_WIDTH - data.MatchedString.Length) / 2;
                        if (data.ColNo - data.ShowColNo - 1 < 0)
                        {
                            data.ShowColNo = data.ColNo;
                            data.ShowLine = data.Line.Substring(0, MAX_WIDTH);
                        }
                        else
                        {
                            if (data.ColNo - data.ShowColNo + MAX_WIDTH > line.Length)
                            {
                                data.ShowColNo = line.Length - MAX_WIDTH;
                                data.ShowLine = data.Line.Substring((int)data.ShowColNo);
                                data.ShowColNo = data.ColNo - data.ShowColNo;
                            }
                            else
                            {
                                data.ShowLine = data.Line.Substring((int)(data.ColNo - data.ShowColNo), MAX_WIDTH);
                            }
                        }
                    }
                }
                else
                {
                    data.ShowColNo = data.ColNo;
                    data.ShowLine = data.Line;
                }

                list.Add(data);
            }

            return;
        }

        /// <summary>
        /// Open File in Specific Editor
        /// </summary>
        /// <param name="data">result data</param>
        public void OpenFileWithEditor(ResultData data)
        {
            try
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                if (!Properties.Settings.Default.UseCustomEditor)
                {
                    p.StartInfo.FileName = string.Format("\"{0}\"", data.FullFileName);
                    p.Start();
                }
                else
                {
                    // col, row, filename
                    p.StartInfo.FileName = string.Format("\"{0}\"", Properties.Settings.Default.CustomEditorPath);
                    p.StartInfo.Arguments = Properties.Settings.Default.CustomEditorArguments.Replace("%file", data.FullFileName.Replace("\\", "\\\\")).Replace("%line", data.RowNo.ToString()).Replace("%column", data.ColNo.ToString());
                    p.Start();
                    if (p.HasExited)
                        p.Kill();
                }
            }
            catch { }
        }
    }
}
