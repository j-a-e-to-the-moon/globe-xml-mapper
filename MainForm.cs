using System;
using System.Collections.Generic;
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

            var btnConvert = new Button
            {
                Text = "XML 변환하기",
                Dock = DockStyle.Fill,
                Height = 38
            };
            btnConvert.Click += BtnConvert_Click;

            layout.Controls.Add(btnOpen, 0, 0);
            layout.Controls.Add(btnNew, 0, 1);
            layout.Controls.Add(btnConvert, 0, 2);

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

        private void BtnConvert_Click(object sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Filter = "Excel 파일 (*.xlsx)|*.xlsx",
                Title = "변환할 서식 파일 선택"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            using var saveDlg = new SaveFileDialog
            {
                Filter = "XML 파일 (*.xml)|*.xml",
                Title = "XML 파일 저장",
                FileName = "GLOBE_OECD.xml"
            };
            if (saveDlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                var globe = new Globe.GlobeOecd
                {
                    Version = "2.0",
                    MessageSpec = new Globe.MessageSpecType(),
                    GlobeBody = new Globe.GlobeBodyType()
                };

                var orchestrator = new MappingOrchestrator();
                var mappingErrors = orchestrator.MapWorkbook(dlg.FileName, globe);

                var xml = XmlExportService.Serialize(globe);
                System.IO.File.WriteAllText(saveDlg.FileName, xml, System.Text.Encoding.UTF8);

                var validationErrors = ValidationUtil.Validate(globe);

                var allErrors = new List<string>();
                if (mappingErrors.Count > 0)
                {
                    allErrors.Add("── 매핑 오류 ──");
                    allErrors.AddRange(mappingErrors);
                    allErrors.Add("");
                }
                if (validationErrors.Count > 0)
                {
                    allErrors.Add("── 검증 오류 (에러코드 기준) ──");
                    allErrors.AddRange(validationErrors);
                }

                var errorsPath = System.IO.Path.ChangeExtension(saveDlg.FileName, ".errors.txt");
                if (allErrors.Count > 0)
                {
                    System.IO.File.WriteAllText(errorsPath,
                        $"[오류 목록] {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}" +
                        $"매핑 오류 {mappingErrors.Count}건 / 검증 오류 {validationErrors.Count}건{Environment.NewLine}{Environment.NewLine}" +
                        string.Join(Environment.NewLine, allErrors),
                        System.Text.Encoding.UTF8);

                    MessageBox.Show(
                        $"XML 생성 완료.\n\n매핑 오류: {mappingErrors.Count}건\n검증 오류: {validationErrors.Count}건\n\n오류 목록: {errorsPath}",
                        "완료 (오류 있음)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    if (System.IO.File.Exists(errorsPath)) System.IO.File.Delete(errorsPath);
                    MessageBox.Show("XML 생성이 완료되었습니다.", "완료",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"XML 변환 오류:\n{ex.Message}", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
