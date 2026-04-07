using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using GlobeMapper.Services;

namespace GlobeMapper
{
    public class MainForm : Form
    {
        private ExcelController _excel;
        private ControlPanelForm _controlPanel;

        private static readonly string TemplatePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "Resources",
            "template.xlsx"
        );

        // ── 시트 라우팅 정의 ─────────────────────────────────────────────
        // Group.xlsx에 포함할 시트 (국가별 계산)
        private static readonly HashSet<string> GroupSheets = new()
        {
            "3.1~3.2.3.2",
            "3.2.4.4(b)",
            "3.3.1~3.4.2",
        };

        // CE_N.xlsx에 포함할 시트 (구성기업별 계산)
        private static readonly HashSet<string> CeSheets = new() { "3.2.4~3.2.4.5" };

        // main.xlsx에 포함할 시트: GroupSheets·CeSheets를 제외한 나머지
        // (1.x, 2, 빈 시트, 향후 3.4.3 추가 시 GroupSheets/CeSheets에 없으면 자동 포함)

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "GIR 2 XML Mapper";
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            var iconPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Resources",
                "app.ico"
            );
            if (File.Exists(iconPath))
                Icon = new System.Drawing.Icon(iconPath);
            ClientSize = new System.Drawing.Size(630, 630);
            BackColor = System.Drawing.Color.FromArgb(30, 30, 32);
            ForeColor = System.Drawing.Color.FromArgb(220, 220, 224);
            Font = new System.Drawing.Font("Segoe UI", 15f);

            // 타이틀 레이블
            var lblTitle = new Label
            {
                Text = "GIR 2 XML Mapper",
                Dock = DockStyle.Top,
                Height = 78,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Segoe UI Semibold", 19f),
                ForeColor = System.Drawing.Color.FromArgb(230, 230, 235),
                BackColor = System.Drawing.Color.FromArgb(22, 22, 24),
            };

            // 구분선
            var divider = new Panel
            {
                Dock = DockStyle.Top,
                Height = 1,
                BackColor = System.Drawing.Color.FromArgb(55, 55, 60),
            };

            // ── 그룹: 템플릿 생성 ─────────────────────────────────────────
            // 헤더 레이블
            var groupHeader = new Label
            {
                Text = "템플릿 생성",
                Dock = DockStyle.Top,
                Height = 48,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Segoe UI", 13f),
                ForeColor = System.Drawing.Color.FromArgb(130, 130, 140),
                BackColor = System.Drawing.Color.FromArgb(32, 32, 36),
                Padding = new Padding(18, 0, 0, 0),
            };
            // 헤더 하단 구분선
            var groupHeaderDiv = new Panel
            {
                Dock = DockStyle.Top,
                Height = 1,
                BackColor = System.Drawing.Color.FromArgb(55, 55, 62),
            };
            // 버튼 영역
            var groupInner = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(18, 15, 18, 18),
                RowCount = 3,
                ColumnCount = 1,
                BackColor = System.Drawing.Color.Transparent,
            };
            groupInner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 3; i++)
                groupInner.RowStyles.Add(new RowStyle(SizeType.Percent, 33.33f));
            groupInner.Controls.Add(
                MakeButton("MNE 생성", BtnCreateTemplate_Click, new Padding(0, 0, 0, 9)),
                0,
                0
            );
            groupInner.Controls.Add(
                MakeButton("합산단위 생성", BtnCreateCountry_Click, new Padding(0, 0, 0, 9)),
                0,
                1
            );
            groupInner.Controls.Add(
                MakeButton("구성기업 생성", BtnCreateCe_Click, new Padding(0, 0, 0, 0)),
                0,
                2
            );

            // 그룹 컨테이너 (테두리 = 배경색)
            var groupBox = new Panel
            {
                Margin = new Padding(0, 0, 0, 24),
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.FromArgb(55, 55, 62), // 테두리 색
                Padding = new Padding(1),
            };
            var groupInside = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.FromArgb(36, 36, 40),
            };
            // DockStyle.Top은 역순: groupInner(Fill) → div → header
            groupInside.Controls.Add(groupInner);
            groupInside.Controls.Add(groupHeaderDiv);
            groupInside.Controls.Add(groupHeader);
            groupBox.Controls.Add(groupInside);

            // ── 전체 레이아웃 ─────────────────────────────────────────────
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(42, 30, 42, 42),
                RowCount = 3,
                ColumnCount = 1,
                BackColor = System.Drawing.Color.Transparent,
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 서식 편집
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // XML 변환하기
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // 템플릿 생성 그룹

            layout.Controls.Add(
                MakeButton("서식 편집", BtnOpen_Click, new Padding(0, 0, 0, 12), primary: true),
                0,
                0
            );
            layout.Controls.Add(
                MakeButton(
                    "XML 변환하기",
                    BtnConvert_Click,
                    new Padding(0, 0, 0, 24),
                    accent: true
                ),
                0,
                1
            );
            layout.Controls.Add(groupBox, 0, 2);

            // 버전 정보 (하단 고정)
            var lblVersion = new Label
            {
                Text = "v1   만료일 2026.6.30   라이선스 DA",
                Dock = DockStyle.Bottom,
                Height = 36,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Segoe UI", 12f),
                ForeColor = System.Drawing.Color.FromArgb(80, 80, 90),
                BackColor = System.Drawing.Color.FromArgb(22, 22, 24),
            };

            // DockStyle.Top은 나중에 추가된 것이 위로 가므로 역순
            Controls.Add(lblVersion);
            Controls.Add(layout);
            Controls.Add(divider);
            Controls.Add(lblTitle);
        }

        private static Button MakeButton(
            string text,
            EventHandler click,
            Padding margin,
            bool primary = false,
            bool accent = false
        )
        {
            var bg =
                accent ? System.Drawing.Color.FromArgb(210, 160, 0)
                : primary ? System.Drawing.Color.FromArgb(200, 90, 15)
                : System.Drawing.Color.FromArgb(44, 44, 48);
            var hover =
                accent ? System.Drawing.Color.FromArgb(225, 175, 10)
                : primary ? System.Drawing.Color.FromArgb(218, 105, 25)
                : System.Drawing.Color.FromArgb(54, 54, 60);
            var fg =
                (accent || primary)
                    ? System.Drawing.Color.White
                    : System.Drawing.Color.FromArgb(210, 210, 215);

            var btn = new Button
            {
                Text = text,
                Dock = DockStyle.Fill,
                Height = 63,
                Margin = margin,
                FlatStyle = FlatStyle.Flat,
                BackColor = bg,
                ForeColor = fg,
                Font = new System.Drawing.Font("Segoe UI", 15f),
                Cursor = Cursors.Hand,
            };
            btn.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(65, 65, 72);
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.MouseOverBackColor = hover;
            btn.Click += click;
            return btn;
        }

        // ── 파일 열기 ────────────────────────────────────────────────────────
        private void BtnOpen_Click(object sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Filter = "Excel 파일 (*.xlsx)|*.xlsx",
                Title = "서식 파일 열기",
            };
            if (dlg.ShowDialog() != DialogResult.OK)
                return;
            OpenExcelAndShowPanel(dlg.FileName);
        }

        // ── 템플릿 생성 ──────────────────────────────────────────────────────
        // 폴더 선택 → 해당 폴더에 main.xlsx 생성 (3.1~3.2.3.2 시트 제외)
        private void BtnCreateTemplate_Click(object sender, EventArgs e)
        {
            if (!CheckTemplate())
                return;

            using var dlg = new FolderBrowserDialog
            {
                Description = "프로젝트 폴더를 선택하세요. 이 폴더에 main.xlsx가 생성됩니다.",
                UseDescriptionForTitle = true,
            };
            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            var savePath = Path.Combine(dlg.SelectedPath, "main.xlsx");
            if (
                File.Exists(savePath)
                && MessageBox.Show(
                    $"이미 main.xlsx가 존재합니다.\n덮어쓰시겠습니까?",
                    "확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                ) != DialogResult.Yes
            )
                return;

            try
            {
                CreateMainFile(TemplatePath, savePath);
                OfferOpenFolder(dlg.SelectedPath, "main.xlsx 생성 완료.");

                _excel = new ExcelController();
                _excel.Open(savePath);
                ShowControlPanel();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"템플릿 생성 오류:\n{ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // ── 국가별 시트 생성 ──────────────────────────────────────────────────
        // 프로젝트 폴더 선택 → 개수 입력 → 1/, 2/, ... 폴더에 3.1~3.2.3.2.xlsx 생성
        private void BtnCreateCountry_Click(object sender, EventArgs e)
        {
            if (!CheckTemplate())
                return;

            using var dlg = new FolderBrowserDialog
            {
                Description = "프로젝트 폴더를 선택하세요 (main.xlsx가 있는 폴더).",
                UseDescriptionForTitle = true,
            };
            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            if (!TryAskCount("생성할 국가 수:", out int count))
                return;

            // 이미 존재하는 번호 폴더 확인
            var existing = Enumerable
                .Range(1, count)
                .Where(i => Directory.Exists(Path.Combine(dlg.SelectedPath, i.ToString())))
                .ToList();
            if (
                existing.Count > 0
                && MessageBox.Show(
                    $"폴더 {string.Join(", ", existing)}이(가) 이미 존재합니다.\n덮어쓰시겠습니까?",
                    "확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                ) != DialogResult.Yes
            )
                return;

            try
            {
                for (int i = 1; i <= count; i++)
                {
                    var folder = Path.Combine(dlg.SelectedPath, i.ToString());
                    Directory.CreateDirectory(folder);
                    var dest = Path.Combine(folder, "Group.xlsx");
                    CreateCountryFile(TemplatePath, dest);
                }

                OfferOpenFolder(dlg.SelectedPath, $"국가별 폴더 {count}개 생성 완료.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"국가별 시트 생성 오류:\n{ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // ── 구성기업 생성 ─────────────────────────────────────────────────────
        // 국가 폴더 선택 → 개수 입력 → 해당 폴더에 CE 파일 생성
        private void BtnCreateCe_Click(object sender, EventArgs e)
        {
            using var dlg = new FolderBrowserDialog
            {
                Description = "국가 폴더를 선택하세요 (숫자 폴더, 예: 1, 2, 3...).",
                UseDescriptionForTitle = true,
            };
            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            if (!TryAskCount("생성할 구성기업 수:", out int count))
                return;

            try
            {
                for (int i = 1; i <= count; i++)
                {
                    var dest = Path.Combine(dlg.SelectedPath, $"CE_{i}.xlsx");
                    CreateCeFile(TemplatePath, dest);
                }

                OfferOpenFolder(dlg.SelectedPath, $"구성기업 파일 {count}개 생성 완료.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"구성기업 생성 오류:\n{ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // ── XML 변환하기 ──────────────────────────────────────────────────────
        private void BtnConvert_Click(object sender, EventArgs e)
        {
            using var terms = new TermsDialog();
            if (terms.ShowDialog(this) != DialogResult.OK)
                return;

            using var dlg = new FolderBrowserDialog
            {
                Description = "프로젝트 폴더를 선택하세요 (main.xlsx가 있는 폴더).",
                UseDescriptionForTitle = true,
            };
            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            using var saveDlg = new SaveFileDialog
            {
                Filter = "XML 파일 (*.xml)|*.xml",
                Title = "XML 파일 저장",
                FileName = "GLOBE_OECD.xml",
                InitialDirectory = dlg.SelectedPath,
            };
            if (saveDlg.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                var globe = new Globe.GlobeOecd
                {
                    Version = "2.0",
                    MessageSpec = new Globe.MessageSpecType(),
                    GlobeBody = new Globe.GlobeBodyType(),
                };

                var orchestrator = new MappingOrchestrator();
                var mappingErrors = orchestrator.MapFolder(dlg.SelectedPath, globe);

                var xml = XmlExportService.Serialize(globe);
                File.WriteAllText(saveDlg.FileName, xml, System.Text.Encoding.UTF8);

                var validationErrors = ValidationUtil.Validate(globe);

                var errorsPath = Path.ChangeExtension(saveDlg.FileName, ".errors.txt");
                if (mappingErrors.Count > 0 || validationErrors.Count > 0)
                {
                    File.WriteAllText(
                        errorsPath,
                        $"[오류 목록] {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}"
                            + $"매핑 오류 {mappingErrors.Count}건 / 검증 오류 {validationErrors.Count}건{Environment.NewLine}{Environment.NewLine}"
                            + string.Join(Environment.NewLine, mappingErrors)
                            + Environment.NewLine
                            + string.Join(Environment.NewLine, validationErrors),
                        System.Text.Encoding.UTF8
                    );
                    MessageBox.Show(
                        $"XML 생성 완료.\n매핑 오류: {mappingErrors.Count}건\n검증 오류: {validationErrors.Count}건\n\n오류 목록: {errorsPath}",
                        "완료 (오류 있음)",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }
                else
                {
                    if (File.Exists(errorsPath))
                        File.Delete(errorsPath);
                    MessageBox.Show(
                        "XML 생성이 완료되었습니다.",
                        "완료",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"XML 변환 오류:\n{ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // ── 파일 생성 헬퍼 ───────────────────────────────────────────────────

        /// <summary>
        /// template → main.xlsx: GroupSheets + CeSheets 제거, 나머지(1.x, 2, 빈 시트, 3.4.3 등) 유지.
        /// </summary>
        private static void CreateMainFile(string templatePath, string savePath)
        {
            File.Copy(templatePath, savePath, overwrite: true);
            using var wb = new XLWorkbook(savePath);
            var toDelete = wb
                .Worksheets.Where(ws => GroupSheets.Contains(ws.Name) || CeSheets.Contains(ws.Name))
                .ToList();
            foreach (var ws in toDelete)
                ws.Delete();
            wb.Save();
        }

        /// <summary>
        /// template → Group.xlsx: GroupSheets에 정의된 시트만 유지.
        /// 향후 3.3.1~3.4.2 시트를 GroupSheets에 추가하면 자동 포함.
        /// </summary>
        private static void CreateCountryFile(string templatePath, string savePath)
        {
            File.Copy(templatePath, savePath, overwrite: true);
            using var wb = new XLWorkbook(savePath);
            var toDelete = wb.Worksheets.Where(ws => !GroupSheets.Contains(ws.Name)).ToList();
            foreach (var ws in toDelete)
                ws.Delete();
            wb.Save();
        }

        /// <summary>
        /// template → CE_N.xlsx: CeSheets에 정의된 시트만 유지.
        /// 향후 CE 섹션 시트를 CeSheets에 추가하면 자동 포함.
        /// </summary>
        private static void CreateCeFile(string templatePath, string savePath)
        {
            if (File.Exists(savePath))
                return; // 이미 있으면 건너뜀
            File.Copy(templatePath, savePath, overwrite: false);
            using var wb = new XLWorkbook(savePath);
            var toDelete = wb.Worksheets.Where(ws => !CeSheets.Contains(ws.Name)).ToList();
            foreach (var ws in toDelete)
                ws.Delete();
            wb.Save();
        }

        // ── 공통 헬퍼 ────────────────────────────────────────────────────────

        private static void OfferOpenFolder(string folderPath, string message)
        {
            var result = MessageBox.Show(
                $"{message}\n\n해당 폴더를 여시겠습니까?",
                "완료",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information
            );
            if (result == DialogResult.Yes)
                System.Diagnostics.Process.Start("explorer.exe", folderPath);
        }

        private bool CheckTemplate()
        {
            if (File.Exists(TemplatePath))
                return true;
            MessageBox.Show(
                "템플릿 파일을 찾을 수 없습니다.\n" + TemplatePath,
                "오류",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return false;
        }

        private static bool TryAskCount(string prompt, out int count)
        {
            count = 0;
            using var f = new Form
            {
                Text = "개수 입력",
                Size = new System.Drawing.Size(540, 270),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false,
                Font = new System.Drawing.Font("Segoe UI", 15f),
            };
            var lbl = new Label
            {
                Text = prompt,
                AutoSize = true,
                Location = new System.Drawing.Point(36, 36),
            };
            var nud = new NumericUpDown
            {
                Location = new System.Drawing.Point(36, 87),
                Width = 165,
                Height = 48,
                Minimum = 1,
                Maximum = 99,
                Value = 1,
                Font = new System.Drawing.Font("Segoe UI", 16f),
            };
            var btnOk = new Button
            {
                Text = "확인",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(330, 84),
                Width = 150,
                Height = 54,
            };
            f.Controls.AddRange(new Control[] { lbl, nud, btnOk });
            f.AcceptButton = btnOk;

            if (f.ShowDialog() != DialogResult.OK)
                return false;
            count = (int)nud.Value;
            return true;
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
                MessageBox.Show(
                    $"파일 열기 오류:\n{ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void ShowControlPanel()
        {
            Hide();
            _controlPanel = new ControlPanelForm(_excel);
            _controlPanel.FormClosed += (s, e) =>
            {
                _excel?.Dispose();
                _excel = null;
                _controlPanel = null;
                Show();
            };
            _controlPanel.Show();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _excel?.Dispose();
            _controlPanel?.Close();
            base.OnFormClosing(e);
        }
    }
}
