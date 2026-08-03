using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using System.ComponentModel;

namespace LibVisioNgSharp.Viewer
{
    public partial class frmMain : Form
    {
        private int pageCount = 0;
        private bool isLoading = false;
        private string filePath = null;
        private List<string> svgs = null;

        BackgroundWorker bgWorker = new BackgroundWorker();

        public frmMain()
        {
            InitializeComponent();

            Label.CheckForIllegalCrossThreadCalls = false;
            WebView2.CheckForIllegalCrossThreadCalls = false;

            this.bgWorker.DoWork += this.BgWorker_DoWork;

            this.webView.EnsureCoreWebView2Async();
        }

        private void BgWorker_DoWork(object? sender, DoWorkEventArgs e)
        {
            this.Convert(this.filePath);
        }

        private void frmMain_Load(object sender, EventArgs e)
        {

        }

        private void Reset()
        {
            this.pageCount = 0;
            this.cboNumber.Items.Clear();
            this.btnFirst.Enabled = this.btnLast.Enabled = this.btnPrevious.Enabled = this.btnNext.Enabled = false;
            this.lblTotal.Text = "0";
            this.lblMessage.Text = "";
            this.lblMessage.ForeColor = Color.Black;
            this.filePath = null;

            this.webView.CoreWebView2.NavigateToString("");
        }

        private void tsmiOpenFile_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.FileName = "";

            DialogResult result = this.openFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                this.Reset();

                string filePath = this.openFileDialog1.FileName;

                this.filePath = filePath;

                this.Text = filePath;

                this.bgWorker.RunWorkerAsync();
            }
        }

        private async Task Convert(string filePath)
        {
            try
            {
                var svgs = Converter.ConvertToSvg(filePath);

                this.svgs = svgs;

                if (svgs != null && svgs.Count > 0)
                {
                    this.pageCount = svgs.Count;

                    this.lblTotal.Text = this.pageCount.ToString();

                    for (int i = 1; i <= this.pageCount; i++)
                    {
                        this.cboNumber.Items.Add(i.ToString());
                    }

                    this.cboNumber.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                this.ShowMessage(ex.Message);
            }
        }

        private void ShowMessage(string message)
        {
            this.lblMessage.Invoke(() =>
            {
                this.lblMessage.Text = message;
            });
        }

        private void cboNumber_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.isLoading)
            {
                return;
            }

            int index = this.cboNumber.SelectedIndex;

            this.ShowHtml(index);
        }

        private async void ShowHtml(int index)
        {
            this.lblMessage.Text = "";

            if (index >= 0 && index < this.pageCount)
            {
                try
                {
                    var svg = this.svgs[index];

                    var html = $"<div style='zoom:1.25'>{svg}</div>";

                    await this.webView.Invoke(async () =>
                    {
                        this.webView.NavigateToString("");

                        this.webView.Source = new Uri("about:blank");

                        string encodedHtml = JsonConvert.SerializeObject(html);
                        string script = "window.document.write(" + encodedHtml + ")";

                        await this.webView.EnsureCoreWebView2Async();
                        await this.webView.ExecuteScriptAsync(script);
                    });
                }
                catch (Exception ex)
                {
                    this.webView.Invoke(() =>
                    {
                        this.webView.CoreWebView2.NavigateToString("");
                    });

                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    this.isLoading = true;

                    this.cboNumber.SelectedIndex = index;

                    this.isLoading = false;

                    this.SetControlStatus(index);
                }
            }
        }

        private void SetControlStatus(int index)
        {
            this.btnFirst.Enabled = index > 0;
            this.btnPrevious.Enabled = index > 0;
            this.btnNext.Enabled = index < this.pageCount - 1;
            this.btnLast.Enabled = index < this.pageCount - 1;
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            this.ShowHtml(0);
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            this.ShowHtml(this.cboNumber.SelectedIndex - 1);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            this.ShowHtml(this.cboNumber.SelectedIndex + 1);
        }

        private void btnLast_Click(object sender, EventArgs e)
        {
            this.ShowHtml(this.pageCount - 1);
        }
    }
}
