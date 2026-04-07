using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using GlobeMapper.Services;

namespace GlobeMapper
{
    public class MainForm : Form
    {
        private ExcelController _excel;
        private ControlPanelForm _controlPanel;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "Globe XML Mapper";
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(300, 180);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20),
                RowCount = 3,
                ColumnCount = 1,
                AutoSize = true
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var btnOpen = new Button
            {
                Text = "파일 열기",
                Dock = DockStyle.Fill,
                Height = 38,
                Margin = new Padding(0, 0, 0, 6)
            };
            btnOpen.Click += BtnOpen_Click;

            var btnNew = new Button
            {
                Text = "새 파일 만들기",
                Dock = DockStyle.Fill,
                Height = 38,
                Margin = new Padding(0, 0, 0, 6)
            };
            btnNew.Click += BtnNew_Click;

            var btnTemplate = new Button
            {
                Text = "템플릿 다운로드",
                Dock = DockStyle.Fill,
                Height = 38
            };
            btnTemplate.Click += BtnTemplate_Click;

            layout.Controls.Add(btnOpen, 0, 0);
            layout.Controls.Add(btnNew, 0, 1);
            layout.Controls.Add(btnTemplate, 0, 2);

            Controls.Add(layout);
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Filter = "Excel 파일 (*.xlsx)|*.xlsx",
                Title = "서식 파일 열기"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            OpenExcelAndShowPanel(dlg.FileName);
        }

        private void BtnNew_Click(object sender, EventArgs e)
        {
            // 템플릿 파일 위치
            var templatePath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Resources", "template.xlsx");

            if (!File.Exists(templatePath))
            {
                // templates 폴더에서 원본 찾기
                templatePath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Resources", "templates", "template.xlsx");
            }

            if (!File.Exists(templatePath))
            {
                MessageBox.Show("템플릿 파일을 찾을 수 없습니다.", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var dlg = new SaveFileDialog
            {
                Filter = "Excel 파일 (*.xlsx)|*.xlsx",
                Title = "새 서식 파일 저장",
                FileName = "GIR_신고서.xlsx"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                _excel = new ExcelController();
                _excel.CreateNew(templatePath, dlg.FileName);
                ShowControlPanel();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일 생성 오류:\n{ex.Message}", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnTemplate_Click(object sender, EventArgs e)
        {
            var templatePath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Resources", "template.xlsx");

            if (!File.Exists(templatePath))
            {
                MessageBox.Show("템플릿 파일을 찾을 수 없습니다.", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var dlg = new SaveFileDialog
            {
                Filter = "Excel 파일 (*.xlsx)|*.xlsx",
                Title = "템플릿 저장 위치 선택",
                FileName = "GIR_template.xlsx"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            File.Copy(templatePath, dlg.FileName, true);
            MessageBox.Show($"템플릿이 저장되었습니다.\n\n{dlg.FileName}",
                "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OpenExcelAndShowPanel(string path)
        {
            try
            {
                _excel = new ExcelController();
                _excel.Open(path);
                ShowControlPanel();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일 열기 오류:\n{ex.Message}", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowControlPanel()
        {
            Hide();

            _controlPanel = new ControlPanelForm(_excel);
            _controlPanel.FormClosed += (s, e) =>
            {
                // Control Panel 닫힘 → 메인화면 복귀
                _excel?.Dispose();
                _excel = null;
                _controlPanel = null;
                Show();
            };
            _controlPanel.Show();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 메인 윈폼 종료 시 엑셀도 함께 종료
            _excel?.Dispose();
            _controlPanel?.Close();
            base.OnFormClosing(e);
        }
    }
}
