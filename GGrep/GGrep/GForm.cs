using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections;
using System.Threading;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using BrightIdeasSoftware;
using GGrep.Instance;
using GGrep.Pattern;
using GGrep.Network;

namespace GGrep
{
    #region Main Class

    /// <summary>
    /// Main Grep Form
    /// </summary>
    public partial class GForm : Form
    {
        #region Members
        public const char SEPERATOR = '	';

        Stopwatch sw = new Stopwatch();

        private SearchOptions option = null;
        public SearchOptions Option { get { return option; } }

        private SearchStatus status = null;
        public SearchStatus Status { get { return status; } set { status = value; } }

        private bool isRunning = false;
        public bool IsRunning { get { return isRunning; } }

        private bool isReplace = false;
        public bool IsReplace { get { return isReplace; } }

        public BackgroundWorker BGW { get { return backgroundWorker; } }

        private bool IsAutoEncoding { get { return GetControlText(cbbEncoding).Trim().ToLower() == "auto"; } }
        #endregion

        #region Constructor
        public GForm(string[] args)
        {
            #region language
            ApplyLanguage();
            InitializeComponent();
            gbFilter.IsCollapsed = Properties.Settings.Default.FilterIsCollapsed;
            SetSearchOptions();
            btnSearch.Text = Properties.Resources.BTN_TEXT_SEARCH;
            btnReplace.Text = Properties.Resources.BTN_TEXT_REPLACE;
            filterToolStripMenuItem.Checked = !Properties.Settings.Default.FilterIsCollapsed;
            tooltipsToolStripMenuItem.Checked = Properties.Settings.Default.TooltipsShown;
            if ("ja-JP".Equals(Utils.GetAppLang()))
            {
                japaneseToolStripMenuItem.Checked = true;
                englishToolStripMenuItem.Checked = false;
                chineseToolStripMenuItem.Checked = false;
            }
            else if ("zh-CN".Equals(Utils.GetAppLang()))
            {
                japaneseToolStripMenuItem.Checked = false;
                englishToolStripMenuItem.Checked = false;
                chineseToolStripMenuItem.Checked = true;
            }
            else
            {
                japaneseToolStripMenuItem.Checked = false;
                englishToolStripMenuItem.Checked = true;
                chineseToolStripMenuItem.Checked = false;
            }
            #endregion

            #region highlight render
            // FIXME
            //this.colResult.Renderer = new HighlightRenderer();
            #endregion

            #region tooltips
            folvResult.CellToolTipGetter = delegate (OLVColumn column, Object rowObject)
            {
                if (rowObject != null && tooltipsToolStripMenuItem.Checked)
                {
                    switch (column.AspectName)
                    {
                        case "FileName":
                            return ((ResultData)rowObject).FullFileName;
                        case "FileEncoding":
                            return ((ResultData)rowObject).FileEncoding;
                        case "Line":
                            return ((ResultData)rowObject).Line.Trim();
                        case "Matched":
                            return ((ResultData)rowObject).MatchedString.Trim();
                        default:
                            return null;
                    }
                }
                else
                    return null;
            };
            #endregion

            #region set window location & size
            // Set window location
            if (Properties.Settings.Default.WindowLocation != null)
                this.Location = Properties.Settings.Default.WindowLocation;

            // Set window size
            if (Properties.Settings.Default.WindowSize != null)
                this.Size = Properties.Settings.Default.WindowSize;
            #endregion

            #region check arguments
            if (args != null && args.Length == 1)
            {
                cbbSearchFolder.SelectedIndex = -1;
                cbbSearchFolder.Text = args[0];
            }
            #endregion

            #region check for updates
            if (Properties.Settings.Default.AutoCheckUpdates)
            {
                DateTime dt = Properties.Settings.Default.LastCheckUpdatesDate.AddDays(Properties.Settings.Default.CheckUpdatesInterval);
                if (dt > DateTime.Now)
                {
                    new UpdateChecker(true).check();
                }
            }
            #endregion
        }
        #endregion

        #region Methods

        #region General

        /// <summary>
        /// Check & Save Search Options
        /// </summary>
        /// <param name="isReplace"></param>
        /// <returns></returns>
        private bool CheckAndSaveSearchOptions(bool isReplace)
        {
            #region check
            // string
            if (string.IsNullOrEmpty(cbbSearchText.Text))
            {
                Utils.ShowMessage(this, 0, Properties.Resources.MSG_ERROR_01);
                return false;
            }

            // folder
            if (string.IsNullOrEmpty(cbbSearchFolder.Text) || !Directory.Exists(cbbSearchFolder.Text))
            {
                Utils.ShowMessage(this, 0, Properties.Resources.MSG_ERROR_02);
                return false;
            }

            // replace
            if (isReplace && string.IsNullOrEmpty(cbbReplaceText.Text))
            {
                if (Utils.ShowMessage(this, 2, Properties.Resources.MSG_WARN_02) != System.Windows.Forms.DialogResult.OK)
                    return false;
            }

            // encoding
            try
            {
                if (!IsAutoEncoding)
                {
                    Encoding e = Encoding.GetEncoding(cbbEncoding.Text);
                }
            }
            catch
            {
                Utils.ShowMessage(this, 0, Properties.Resources.MSG_ERROR_03);
                return false;
            }

            // regex text
            if (cbRegex.Checked)
            {
                try
                {
                    Regex.IsMatch("", cbbSearchText.Text);
                }
                catch
                {
                    Utils.ShowMessage(this, 0, Properties.Resources.MSG_ERROR_04);
                    cbbSearchText.Focus();
                    return false;
                }

                try
                {
                    Regex.IsMatch("", cbbReplaceText.Text);
                }
                catch
                {
                    Utils.ShowMessage(this, 0, Properties.Resources.MSG_ERROR_04);
                    cbbReplaceText.Focus();
                    return false;
                }
            }

            #endregion

            #region save
            option = new SearchOptions();

            option.SearchString = cbbSearchText.Text;
            ManagerCombox(cbbSearchText, option.SearchString);

            option.SearchFolder = cbbSearchFolder.Text;
            ManagerCombox(cbbSearchFolder, option.SearchFolder);

            option.ReplaceString = cbbReplaceText.Text;
            ManagerCombox(cbbReplaceText, option.ReplaceString);

            option.IncludeSubFolders = cbIncludeSubFolders.Checked;
            option.IncludeHiddenFolders = cbIncludeHiddenFolder.Checked;

            option.IsRegex = cbRegex.Checked;
            if (cbRegex.Checked)
            {
                option.Multiline = cbMultiline.Checked;
            }
            else
            {
                option.IsCaseSensitive = cbCaseSensitive.Checked;
                option.IsSearchOnWords = cbSearchOnWords.Checked;
            }

            option.Encoding = cbbEncoding.Text;
            ManagerCombox(cbbEncoding, option.Encoding);

            option.MatchFiles = cbbFileMatch.Text;
            ManagerCombox(cbbFileMatch, option.MatchFiles);

            option.NotMatchFiles = cbbFileNotMatch.Text;
            ManagerCombox(cbbFileNotMatch, option.NotMatchFiles);

            option.MatchFolders = cbbFolderMatch.Text;
            ManagerCombox(cbbFolderMatch, option.MatchFolders);

            option.NotMatchFolders = cbbFolderNotMatch.Text;
            ManagerCombox(cbbFolderNotMatch, option.NotMatchFolders);

            #endregion

            return true;
        }

        /// <summary>
        /// Manager ComboBox
        /// </summary>
        /// <param name="cb">ComboBox</param>
        /// <param name="value">Selected Value</param>
        private void ManagerCombox(ComboBox cb, string value)
        {
            if (cb.Items.Contains(value))
            {
                cb.Items.Remove(value);
            }
            cb.Items.Insert(0, value);
            cb.SelectedIndex = 0;
        }

        /// <summary>
        /// generate string like 'xxx|xxx|xxx|xxx'
        /// </summary>
        /// <param name="cb">select box</param>
        private string AppendSearchOptions(ComboBox cb)
        {
            if (cb.Items.Count <= 0)
                return "";
            string s = "";

            // saved option max count is 20
            int count = cb.Items.Count > 20 ? 20 : cb.Items.Count;

            for (int i = 0; i < count; i++)
            {
                s += SEPERATOR + cb.Items[i].ToString();
            }
            return s.Substring(1);
        }

        /// <summary>
        /// Save Search Options
        /// </summary>
        private void SaveSearchOptions()
        {
            Properties.Settings.Default.SearchString = AppendSearchOptions(cbbSearchText);
            Properties.Settings.Default.SearchFolder = AppendSearchOptions(cbbSearchFolder);
            Properties.Settings.Default.ReplaceString = AppendSearchOptions(cbbReplaceText);
            Properties.Settings.Default.Regex = option.IsRegex;
            Properties.Settings.Default.CaseSensitive = option.IsCaseSensitive;
            Properties.Settings.Default.SearchOnWords = option.IsSearchOnWords;
            Properties.Settings.Default.IncludeSubFolders = option.IncludeSubFolders;
            Properties.Settings.Default.IncludeHiddenFolders = option.IncludeHiddenFolders;
            Properties.Settings.Default.Multiline = option.Multiline;
            Properties.Settings.Default.Encoding = AppendSearchOptions(cbbEncoding);
            Properties.Settings.Default.FileMatch = AppendSearchOptions(cbbFileMatch);
            Properties.Settings.Default.FileNotMatch = AppendSearchOptions(cbbFileNotMatch);
            Properties.Settings.Default.FolderMatch = AppendSearchOptions(cbbFolderMatch);
            Properties.Settings.Default.FolderNotMatch = AppendSearchOptions(cbbFolderNotMatch);
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Set Search Options
        /// </summary>
        private void SetSearchOptions()
        {
            if (!string.IsNullOrEmpty(Properties.Settings.Default.SearchString))
            {
                cbbSearchText.Items.Clear();
                foreach (string s in Properties.Settings.Default.SearchString.Split(SEPERATOR))
                {
                    cbbSearchText.Items.Add(s);
                }
                cbbSearchText.SelectedIndex = 0;
            }
            if (!string.IsNullOrEmpty(Properties.Settings.Default.SearchFolder))
            {
                cbbSearchFolder.Items.Clear();
                foreach (string s in Properties.Settings.Default.SearchFolder.Split(SEPERATOR))
                {
                    cbbSearchFolder.Items.Add(s);
                }
                cbbSearchFolder.SelectedIndex = 0;
            }
            if (!string.IsNullOrEmpty(Properties.Settings.Default.ReplaceString))
            {
                cbbReplaceText.Items.Clear();
                foreach (string s in Properties.Settings.Default.ReplaceString.Split(SEPERATOR))
                {
                    cbbReplaceText.Items.Add(s);
                }
                cbbReplaceText.SelectedIndex = 0;
            }
            cbRegex.Checked = Properties.Settings.Default.Regex;
            cbCaseSensitive.Checked = Properties.Settings.Default.CaseSensitive;
            cbSearchOnWords.Checked = Properties.Settings.Default.SearchOnWords;
            cbSearchOnWords.Enabled = !cbRegex.Checked;
            cbCaseSensitive.Enabled = !cbRegex.Checked;
            cbMultiline.Enabled = cbRegex.Checked;

            cbIncludeSubFolders.Checked = Properties.Settings.Default.IncludeSubFolders;
            cbIncludeHiddenFolder.Checked = Properties.Settings.Default.IncludeHiddenFolders;
            cbIncludeHiddenFolder.Enabled = cbIncludeSubFolders.Checked;
            cbMultiline.Checked = Properties.Settings.Default.Multiline;

            if (!string.IsNullOrEmpty(Properties.Settings.Default.Encoding))
            {
                cbbEncoding.Items.Clear();
                foreach (string s in Properties.Settings.Default.Encoding.Split(SEPERATOR))
                {
                    cbbEncoding.Items.Add(s);
                }
                cbbEncoding.SelectedIndex = 0;
            }
            if (!string.IsNullOrEmpty(Properties.Settings.Default.FileMatch))
            {
                cbbFileMatch.Items.Clear();
                foreach (string s in Properties.Settings.Default.FileMatch.Split(SEPERATOR))
                {
                    cbbFileMatch.Items.Add(s);
                }
                cbbFileMatch.SelectedIndex = 0;
            }
            if (!string.IsNullOrEmpty(Properties.Settings.Default.FileNotMatch))
            {
                cbbFileNotMatch.Items.Clear();
                foreach (string s in Properties.Settings.Default.FileNotMatch.Split(SEPERATOR))
                {
                    cbbFileNotMatch.Items.Add(s);
                }
                cbbFileNotMatch.SelectedIndex = 0;
            }
            if (!string.IsNullOrEmpty(Properties.Settings.Default.FolderMatch))
            {
                cbbFolderMatch.Items.Clear();
                foreach (string s in Properties.Settings.Default.FolderMatch.Split(SEPERATOR))
                {
                    cbbFolderMatch.Items.Add(s);
                }
                cbbFolderMatch.SelectedIndex = 0;
            }
            if (!string.IsNullOrEmpty(Properties.Settings.Default.FolderNotMatch))
            {
                cbbFolderNotMatch.Items.Clear();
                foreach (string s in Properties.Settings.Default.FolderNotMatch.Split(SEPERATOR))
                {
                    cbbFolderNotMatch.Items.Add(s);
                }
                cbbFolderNotMatch.SelectedIndex = 0;
            }

        }

        /// <summary>
        /// Set language
        /// </summary>
        /// <param name="lang"></param>
        private void SetLanguage(string lang)
        {
            Properties.Settings.Default.Language = lang;
            Properties.Settings.Default.Save();
            Application.Restart();
        }

        private void ApplyLanguage()
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(Utils.GetAppLang());
            Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo(Utils.GetAppLang());
        }
        #endregion

        #region CallBack
        delegate void SetControlTextCallback(Control ctl, string text);
        private void SetControlText(Control ctl, string text)
        {
            if (ctl.InvokeRequired)
            {
                SetControlTextCallback d = new SetControlTextCallback(SetControlText);
                this.Invoke(d, new object[] { ctl, text });
            }
            else
            {
                ctl.Text = text;
            }
        }

        delegate void SetControlActiveCallback(Control ctl, bool enable);
        private void SetControlActive(Control ctl, bool enable)
        {
            if (ctl.InvokeRequired)
            {
                SetControlActiveCallback d = new SetControlActiveCallback(SetControlActive);
                this.Invoke(d, new object[] { ctl, enable });
            }
            else
            {
                ctl.Enabled = enable;
            }
        }

        delegate void SetToolStripMenuItemActiveCallback(ToolStripMenuItem ctl, bool enable);
        private void SetToolStripMenuItemActive(ToolStripMenuItem ctl, bool enable)
        {
            if (this.InvokeRequired)
            {
                SetToolStripMenuItemActiveCallback d = new SetToolStripMenuItemActiveCallback(SetToolStripMenuItemActive);
                this.Invoke(d, new object[] { ctl, enable });
            }
            else
            {
                ctl.Enabled = enable;
            }
        }

        delegate string GetControlTextCallback(Control ctl);
        private string GetControlText(Control ctl)
        {
            if (ctl.InvokeRequired)
            {
                GetControlTextCallback d = new GetControlTextCallback(GetControlText);
                return (string)this.Invoke(d, new object[] { ctl });
            }
            else
            {
                return ctl.Text;
            }
        }

        delegate void AddRowsCallback(ObjectListView list, ArrayList dataList, SearchStatus status);
        private void AddRows(ObjectListView list, ArrayList dataList, SearchStatus status)
        {
            if (list.InvokeRequired)
            {
                AddRowsCallback d = new AddRowsCallback(AddRows);
                this.Invoke(d, new object[] { list, dataList, status });
            }
            else
            {
                list.AddObjects(dataList);
            }
        }

        delegate void SetToolStripLabelCallback(ToolStripStatusLabel ctl, string text);
        private void SetToolStripLabel(ToolStripStatusLabel ctl, string text)
        {
            if (InvokeRequired)
            {
                SetToolStripLabelCallback d = new SetToolStripLabelCallback(SetToolStripLabel);
                this.Invoke(d, new object[] { ctl, text });
            }
            else
            {
                ctl.Text = text;
            }
        }
        #endregion

        #region Output Csv
        /// <summary>
        /// Save as CSV
        /// </summary>
        /// <returns></returns>
        private bool SaveCsvFile()
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = Properties.Resources.FILE_FILTER;
                sfd.FileName = Properties.Resources.FILE_NEM_NAME;
                sfd.RestoreDirectory = true;
                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    // save
                    using (StreamWriter sw = new StreamWriter(sfd.FileName, false, Encoding.GetEncoding("SJIS")))
                    {
                        #region write header
                        sw.Write("#");
                        if (colFileName.IsVisible)
                            sw.Write(",File");
                        if (colEncoding.IsVisible)
                            sw.Write(",Encoding");
                        if (colRowNo.IsVisible)
                            sw.Write(",Row");
                        if (colColNo.IsVisible)
                            sw.Write(",Column");
                        if (colResult.IsVisible)
                            sw.Write(",Result");
                        if (colMatched.IsVisible)
                            sw.Write(",Matched");
                        sw.Write(System.Environment.NewLine);
                        #endregion

                        int index = 1;
                        foreach (ResultData data in folvResult.Objects)
                        {
                            sw.WriteLine(ToCsvLine(data));
                            if (index++ % 1000 == 0)
                                sw.Flush();
                        }
                        sw.Close();
                    }
                    Utils.ShowMessage(this, 1, Properties.Resources.MSG_SAVED);
                }

            }
            catch (Exception e)
            {
                Utils.ShowMessage(this, 0, e.ToString());
                return false;
            }
            return true;
        }

        /// <summary>
        /// Make Csv Line
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private string ToCsvLine(ResultData data)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(data.No);
            if (colFileName.IsVisible)
            {
                sb.Append(",");
                sb.Append(data.FullFileName);
            }
            if (colEncoding.IsVisible)
            {
                sb.Append(",");
                sb.Append(data.FileEncoding);
            }
            if (colRowNo.IsVisible)
            {
                sb.Append(",");
                sb.Append(data.RowNo);
            }
            if (colColNo.IsVisible)
            {
                sb.Append(",");
                sb.Append(data.ColNo);
            }
            if (colResult.IsVisible)
            {
                sb.Append(",");
                sb.Append("\"");
                sb.Append(data.Line.Replace("\"", "\"\""));
                sb.Append("\"");
            }
            if (colMatched.IsVisible)
            {
                sb.Append(",");
                sb.Append("\"");
                sb.Append(data.MatchedString.Replace("\"", "\"\""));
                sb.Append("\"");
            }

            return sb.ToString();
        }
        #endregion

        #endregion

        #region Events

        #region button, combobox, checkbox and etc.
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (!string.IsNullOrEmpty(cbbSearchFolder.Text) && Directory.Exists(cbbSearchFolder.Text))
            {
                fbd.SelectedPath = cbbSearchFolder.Text;
            }
            switch (fbd.ShowDialog(this))
            {
                case DialogResult.OK:
                    cbbSearchFolder.Text = fbd.SelectedPath;
                    break;
                default:
                    break;
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (!isRunning && !backgroundWorker.IsBusy)
            {
                // start
                if (!CheckAndSaveSearchOptions(false))
                    return;

                SaveSearchOptions();

                isReplace = false;
                backgroundWorker.RunWorkerAsync();
            }
            else
            {
                // stop
                isRunning = false;
                backgroundWorker.CancelAsync();
            }
        }


        private void btnReplace_Click(object sender, EventArgs e)
        {
            if (!isRunning && !backgroundWorker.IsBusy)
            {
                // start
                if (!CheckAndSaveSearchOptions(true))
                    return;

                SaveSearchOptions();

                isReplace = true;
                backgroundWorker.RunWorkerAsync();
            }
            else
            {
                // stop
                isRunning = false;
                backgroundWorker.CancelAsync();
            }
        }

        private void cbRegex_CheckedChanged(object sender, EventArgs e)
        {
            cbMultiline.Enabled = cbRegex.Checked;
            cbSearchOnWords.Enabled = !cbRegex.Checked;
            cbCaseSensitive.Enabled = !cbRegex.Checked;
        }

        private void cbIncludeSubFolders_CheckedChanged(object sender, EventArgs e)
        {
            cbIncludeHiddenFolder.Enabled = cbIncludeSubFolders.Checked;
        }

        private void GForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (isRunning || backgroundWorker.IsBusy)
            {
                if (Utils.ShowMessage(this, 2, Properties.Resources.MSG_WARN_01) != DialogResult.OK)
                {
                    e.Cancel = true;
                    return;
                }
                // stop
                isRunning = false;
                backgroundWorker.CancelAsync();
            }

            // Copy window location to app settings
            Properties.Settings.Default.WindowLocation = this.Location;

            // Copy window size to app settings
            if (this.WindowState == FormWindowState.Normal)
                Properties.Settings.Default.WindowSize = this.Size;
            else
                Properties.Settings.Default.WindowSize = this.RestoreBounds.Size;

            Properties.Settings.Default.FilterIsCollapsed = gbFilter.IsCollapsed;
            Properties.Settings.Default.TooltipsShown = tooltipsToolStripMenuItem.Checked;
            Properties.Settings.Default.Save();
        }

        private void folvResult_DoubleClick(object sender, EventArgs e)
        {
            if (((FastObjectListView)sender).SelectedObject != null)
            {
                ResultData data = (ResultData)((FastObjectListView)sender).SelectedObject;
                PatternFactory.GetPattern(this, data.FileName).OpenFileWithEditor(data);
            }
        }

        #endregion

        #region menu
        private void saveAsCsvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (folvResult.Items.Count <= 0)
            {
                Utils.ShowMessage(this, 0, Properties.Resources.MSG_ERROR_05);
                return;
            }

            SaveCsvFile();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void filterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            gbFilter.IsCollapsed = !gbFilter.IsCollapsed;
            filterToolStripMenuItem.Checked = !gbFilter.IsCollapsed;
        }

        private void optionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Optoins o = new Optoins();
            o.ShowDialog(this);
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About a = new About();
            a.ShowDialog(this);
        }

        private void checkForUpdateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new UpdateChecker(false).check();
        }

        private void japaneseChangeToolStripMenuItem_Clicked(object sender, EventArgs e)
        {
            if (!japaneseToolStripMenuItem.Checked)
            {
                japaneseToolStripMenuItem.Checked = !japaneseToolStripMenuItem.Checked;
                englishToolStripMenuItem.Checked = !englishToolStripMenuItem.Checked;
                chineseToolStripMenuItem.Checked = !chineseToolStripMenuItem.Checked;
                SetLanguage("ja-JP");
            }
        }

        private void englishChangeToolStripMenuItem_Clicked(object sender, EventArgs e)
        {
            if (!englishToolStripMenuItem.Checked)
            {
                japaneseToolStripMenuItem.Checked = !japaneseToolStripMenuItem.Checked;
                englishToolStripMenuItem.Checked = !englishToolStripMenuItem.Checked;
                chineseToolStripMenuItem.Checked = !chineseToolStripMenuItem.Checked;
                SetLanguage("en-US");
            }
        }

        private void chineseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!chineseToolStripMenuItem.Checked)
            {
                japaneseToolStripMenuItem.Checked = !japaneseToolStripMenuItem.Checked;
                englishToolStripMenuItem.Checked = !englishToolStripMenuItem.Checked;
                chineseToolStripMenuItem.Checked = !chineseToolStripMenuItem.Checked;
                SetLanguage("zh-CN");
            }
        }

        #endregion

        #region backgroundworker
        private void DoGrep(object sender, DoWorkEventArgs e)
        {
            ApplyLanguage();
            isRunning = true;
            status = new SearchStatus();
            if (isReplace)
            {
                SetControlText(btnReplace, Properties.Resources.BTN_TEXT_CANCEL);
                SetControlActive(btnSearch, false);
            }
            else
            {
                SetControlText(btnSearch, Properties.Resources.BTN_TEXT_CANCEL);
                SetControlActive(btnReplace, false);
            }
            folvResult.ClearObjects();
            SetToolStripMenuItemActive(saveAsCsvToolStripMenuItem, false);
            SetControlActive(cbbSearchText, false);
            SetControlActive(cbbSearchFolder, false);
            SetControlActive(cbbReplaceText, false);
            SetControlActive(btnBrowse, false);
            SetControlActive(gbFilter, false);
            SetControlActive(menuStripMain, false);
            sw.Reset();
            sw.Start();
            SetToolStripLabel(statusLabel, Properties.Resources.MSG_STATUS_STARTED);

            ArrayList fileList = new ArrayList();
            Utils.AnalyzeDirectory(option.SearchFolder, option, fileList, backgroundWorker, e);

            if (fileList.Count > 0)
            {
                status.Total = fileList.Count;

                // FIXME

                TextMatchFilter filter = TextMatchFilter.Contains(this.folvResult, option.SearchString);
                this.folvResult.ModelFilter = filter;
                this.folvResult.DefaultRenderer = new HighlightRenderer(filter);


                foreach (string file in fileList)
                {
                    if (backgroundWorker.CancellationPending)
                    {
                        e.Cancel = true;
                        break;
                    }

                    ArrayList list = PatternFactory.GetPattern(this, file).SearchInFile(file, e);
                    status.Finished++;
                    if (list.Count > 0)
                        AddRows(folvResult, list, status);
                    backgroundWorker.ReportProgress(status.Progress);
                }
            }
        }

        private void GrepProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (isRunning)
            {
                toolStripProgressBar.Value = e.ProgressPercentage;
                SetToolStripLabel(statusLabel, string.Format(Properties.Resources.MSG_STATUS_RUNNING, status.Finished, status.Total, status.Hit));
            }
        }

        private void GrepCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            sw.Stop();
            toolStripProgressBar.Value = 0;
            if (e.Error != null)
            {
                Utils.ShowMessage(this, 0, e.Error.Message);
                SetToolStripLabel(statusLabel, e.Error.Message);
            }
            else
            {
                if (isReplace)
                {
                    SetToolStripLabel(statusLabel, string.Format((e.Cancelled ? Properties.Resources.MSG_STATUS_REPLACE_CANCELED : Properties.Resources.MSG_STATUS_REPLACE_FINISHED), status.Hit, status.MatchedFiles));
                }
                else
                {
                    SetToolStripLabel(statusLabel, string.Format((e.Cancelled ? Properties.Resources.MSG_STATUS_SEARCH_CANCELED : Properties.Resources.MSG_STATUS_SEARCH_FINISHED), status.Hit, sw.ElapsedMilliseconds));
                }
            }
            SetControlActive(menuStripMain, true);
            SetControlActive(cbbSearchText, true);
            SetControlActive(cbbSearchFolder, true);
            SetControlActive(cbbReplaceText, true);
            SetControlActive(btnBrowse, true);
            SetControlActive(gbFilter, true);
            if (status.Hit > 0)
                SetToolStripMenuItemActive(saveAsCsvToolStripMenuItem, true);

            if (isReplace)
            {
                SetControlText(btnReplace, Properties.Resources.BTN_TEXT_REPLACE);
                SetControlActive(btnSearch, true);
            }
            else
            {
                SetControlText(btnSearch, Properties.Resources.BTN_TEXT_SEARCH);
                SetControlActive(btnReplace, true);
            }
            isRunning = false;
        }
        #endregion

        #region drag
        private void SearchFolder_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All; // Okay
            else
                e.Effect = DragDropEffects.None; // Unknown data, ignore it
        }

        private void SearchFolder_DragDrop(object sender, DragEventArgs e)
        {
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (fileList != null && fileList.Length > 0)
            {
                if (Directory.Exists(fileList[0]))
                {
                    cbbSearchFolder.Text = fileList[0];
                }
                else if (File.Exists(fileList[0]))
                {
                    cbbSearchFolder.Text = Path.GetDirectoryName(fileList[0]);
                }
            }
        }

        private void SearchText_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.Text))
                e.Effect = DragDropEffects.All; // Okay
            else
                e.Effect = DragDropEffects.None; // Unknown data, ignore it
        }

        private void SearchText_DragDrop(object sender, DragEventArgs e)
        {
            cbbSearchText.Text = (string)e.Data.GetData(DataFormats.Text, false);
        }
        #endregion
        #endregion
    }

    #endregion

    #region internal class

    /// <summary>
    /// Highlight Matched String
    /// </summary>
    internal class HighlightRenderer : HighlightTextRenderer
    {
        internal const TextFormatFlags NormalTextFormatFlags = TextFormatFlags.NoPrefix |
                                                               TextFormatFlags.EndEllipsis |
                                                               TextFormatFlags.PreserveGraphicsTranslateTransform;
        public HighlightRenderer(TextMatchFilter filter)
        {
            this.Filter = filter;
            this.FramePen = null;
        }

        protected override void DrawGdiTextHighlighting(Graphics g, Rectangle r, string txt)
        {
            if ((this.IsItemSelected && this.ListView.Focused))
            {
                // if row is selected, skip
            }
            else
            {
                // TextRenderer puts horizontal padding around the strings, so we need to take
                // that into account when measuring strings
                const int paddingAdjustment = 6;

                // Cache the font
                Font f = this.Font;

                foreach (CharacterRange range in this.Filter.FindAllMatchedRanges(txt))
                {
                    // if not the column
                    if (range.First + 1 != ((int)((ResultData)this.RowObject).ShowColNo))
                        continue;

                    // Measure the text that comes before our substring
                    Size precedingTextSize = Size.Empty;
                    if (range.First > 0)
                    {
                        string precedingText = txt.Substring(0, range.First);
                        precedingTextSize = TextRenderer.MeasureText(g, precedingText, f, r.Size, NormalTextFormatFlags);
                        precedingTextSize.Width -= paddingAdjustment;
                    }

                    // Measure the length of our substring (may be different each time due to case differences)
                    string highlightText = txt.Substring(range.First, range.Length);
                    Size textToHighlightSize = TextRenderer.MeasureText(g, highlightText, f, r.Size, NormalTextFormatFlags);
                    textToHighlightSize.Width -= paddingAdjustment;

                    float textToHighlightLeft = r.X + precedingTextSize.Width + 1;
                    float textToHighlightTop = this.AlignVertically(r, textToHighlightSize.Height);

                    // Draw a filled frame around our substring
                    this.DrawSubstringFrame(g, textToHighlightLeft, textToHighlightTop, textToHighlightSize.Width, textToHighlightSize.Height);
                }
            }
        }

    }


    #endregion
}