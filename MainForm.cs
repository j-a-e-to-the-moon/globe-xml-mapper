using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using GlobeMapper.Services;

namespace GlobeMapper
{
    public class MainForm : Form
    {
        private static readonly string TemplatePath = ResolveTemplatePath();

        private static string ResolveTemplatePath()
        {
            var resDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
            try
            {
                var cfgPath = Path.Combine(resDir, "config.json");
                if (File.Exists(cfgPath))
                {
                    var json = File.ReadAllText(cfgPath, System.Text.Encoding.UTF8);
                    using var doc = System.Text.Json.JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("templateFileName", out var prop))
                    {
                        var name = prop.GetString();
                        if (!string.IsNullOrWhiteSpace(name))
                            return Path.Combine(resDir, name);
                    }
                }
            }
            catch { }
            return Path.Combine(resDir, "main_template_newest.xlsx");
        }

        // ── 색상 상수 ──────────────────────────────────────────────────────
        private static readonly Color BG          = Color.FromArgb(28, 28, 30);
        private static readonly Color BG_CARD     = Color.FromArgb(44, 44, 48);
        private static readonly Color BG_HOVER    = Color.FromArgb(54, 54, 60);
        private static readonly Color FG          = Color.FromArgb(230, 230, 235);
        private static readonly Color FG_SUB      = Color.FromArgb(150, 150, 158);
        private static readonly Color FG_DIM      = Color.FromArgb(100, 100, 108);
        private static readonly Color ACCENT      = Color.FromArgb(200, 90, 15);   // XML 변환: 강조
        private static readonly Color ACCENT_HOV  = Color.FromArgb(218, 105, 25);
        private static readonly Color NUM_DIM     = Color.FromArgb(180, 180, 185);
        private static readonly Color NUM_ACCENT  = Color.FromArgb(255, 180, 120);

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text            = "GIR 2 XML Mapper";
            AutoScaleMode   = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox     = false;
            StartPosition   = FormStartPosition.CenterScreen;
            ClientSize      = new Size(640, 420);
            BackColor       = BG;
            ForeColor       = FG;
            Font            = new Font("Segoe UI", 10f);

            var iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "app.ico");
            if (File.Exists(iconPath)) Icon = new Icon(iconPath);

            // ── 타이틀 (컴팩트) ──────────────────────────────────────────
            // 타이틀은 영문이라 Segoe UI 유지
            var title = new Label
            {
                Text      = "GIR 2 XML Mapper",
                Dock      = DockStyle.Top, Height = 44,
                TextAlign = ContentAlignment.MiddleLeft,
                Font      = new Font("Segoe UI Semibold", 13f),
                ForeColor = FG,
                Padding   = new Padding(22, 0, 0, 0),
            };

            // ── 버전 (하단, 플랫) ────────────────────────────────────────
            var ver = new Label
            {
                Text      = "v1  ·  만료일 2026.6.30  ·  라이선스 DA",
                Dock      = DockStyle.Bottom, Height = 30,
                TextAlign = ContentAlignment.MiddleCenter,
                Font      = new Font("Segoe UI", 8.5f),
                ForeColor = FG_DIM,
            };

            // ── 중앙 스텝 영역 ───────────────────────────────────────────
            var content = new Panel
            {
                Dock       = DockStyle.Fill,
                BackColor  = Color.Transparent,
                Padding    = new Padding(22, 6, 22, 10),
            };

            var card1 = MakeStepCard(1, "템플릿 다운로드", "빈 서식 파일 내려받기",         false, BtnCreateMne_Click);
            var card2 = MakeStepCard(2, "XML 변환",        "작성된 xlsx를 GIR XML로 변환", true,  BtnConvert_Click);

            content.Controls.Add(card1);
            content.Controls.Add(card2);

            content.Resize += (s, e) =>
            {
                var x  = content.Padding.Left;
                var w  = content.ClientSize.Width - content.Padding.Left - content.Padding.Right;
                const int cardH = 90;
                const int gap   = 14;
                var totalH = cardH * 2 + gap;
                var y0     = content.Padding.Top + Math.Max(0, (content.ClientSize.Height - content.Padding.Top - content.Padding.Bottom - totalH) / 2);
                card1.SetBounds(x, y0,               w, cardH);
                card2.SetBounds(x, y0 + cardH + gap, w, cardH);
            };

            Controls.Add(content);
            Controls.Add(ver);
            Controls.Add(title);
        }

        /// <summary>
        /// 스텝 카드: [번호] [제목 / 설명] 레이아웃. 카드 전체가 클릭 영역.
        /// accent=true면 강조색(③ XML 변환용).
        /// </summary>
        private Control MakeStepCard(int stepNumber, string title, string desc, bool accent, EventHandler click)
        {
            var baseBg   = accent ? ACCENT     : BG_CARD;
            var hoverBg  = accent ? ACCENT_HOV : BG_HOVER;
            var titleFg  = accent ? Color.White : FG;
            var descFg   = accent ? Color.FromArgb(255, 230, 210) : FG_SUB;
            var numFg    = accent ? NUM_ACCENT : NUM_DIM;

            // 레이아웃 상수
            const int NUM_LEFT   = 8;
            const int NUM_W      = 48;
            const int TEXT_LEFT  = NUM_LEFT + NUM_W + 8; // 64
            const int ARROW_W    = 36;

            var card = new Panel
            {
                BackColor = baseBg,
                Cursor    = Cursors.Hand,
            };

            // AutoEllipsis + GDI 렌더링 조합이 한글 Bold에서 일부 자소 획을 깎는 증상 회피:
            //  - AutoEllipsis = false (잘림 기능 제거; 폭은 Reflow에서 충분히 확보)
            //  - UseCompatibleTextRendering = true (GDI+ 사용, 한글 렌더링 더 안정)
            //  - 초기 Width 100 대신 충분히 큰 값으로 시작
            const int CARD_H = 90;

            var lblNum = new Label
            {
                Text      = stepNumber.ToString(),
                Font      = new Font("Segoe UI Light", 30f),
                ForeColor = numFg,
                BackColor = baseBg,
                Bounds    = new Rectangle(NUM_LEFT, 0, NUM_W, CARD_H),
                TextAlign = ContentAlignment.MiddleCenter,
                UseCompatibleTextRendering = true,
            };

            var lblTitle = new Label
            {
                Text         = title,
                Font         = new Font("Malgun Gothic", 13f, FontStyle.Bold),
                ForeColor    = titleFg,
                BackColor    = baseBg,
                Bounds       = new Rectangle(TEXT_LEFT, 18, 400, 28),
                TextAlign    = ContentAlignment.MiddleLeft,
                AutoEllipsis = false,
                UseCompatibleTextRendering = true,
            };

            var lblDesc = new Label
            {
                Text         = desc,
                Font         = new Font("Malgun Gothic", 10f),
                ForeColor    = descFg,
                BackColor    = baseBg,
                Bounds       = new Rectangle(TEXT_LEFT, 48, 400, 24),
                TextAlign    = ContentAlignment.MiddleLeft,
                AutoEllipsis = false,
                UseCompatibleTextRendering = true,
            };

            var lblArrow = new Label
            {
                Text      = "›",
                Font      = new Font("Segoe UI", 22f),
                ForeColor = accent ? Color.White : FG_SUB,
                BackColor = baseBg,
                Bounds    = new Rectangle(0, 0, ARROW_W, CARD_H),
                TextAlign = ContentAlignment.MiddleCenter,
                UseCompatibleTextRendering = true,
            };

            card.Controls.Add(lblNum);
            card.Controls.Add(lblTitle);
            card.Controls.Add(lblDesc);
            card.Controls.Add(lblArrow);

            // 카드 리사이즈 시 제목/설명/화살표 위치 조정
            void Reflow()
            {
                var w = card.ClientSize.Width;
                int textW = Math.Max(10, w - TEXT_LEFT - ARROW_W - 6);
                lblTitle.Width  = textW;
                lblDesc.Width   = textW;
                lblArrow.Left   = w - ARROW_W - 4;
            }
            card.Resize += (s, e) => Reflow();
            card.HandleCreated += (s, e) => Reflow();

            // 호버 효과 — 카드 및 자식 컨트롤에서 진입/이탈 감지
            // 라벨 BackColor도 동기화 (Transparent 대신 고정색 쓰므로)
            void SetHover(bool on)
            {
                var bg = on ? hoverBg : baseBg;
                card.BackColor    = bg;
                lblNum.BackColor  = bg;
                lblTitle.BackColor = bg;
                lblDesc.BackColor  = bg;
                lblArrow.BackColor = bg;
            }
            card.MouseEnter += (s, e) => SetHover(true);
            card.MouseLeave += (s, e) =>
            {
                if (!card.ClientRectangle.Contains(card.PointToClient(Cursor.Position)))
                    SetHover(false);
            };
            foreach (Control c in new Control[] { lblNum, lblTitle, lblDesc, lblArrow })
            {
                c.MouseEnter += (s, e) => SetHover(true);
                c.MouseLeave += (s, e) =>
                {
                    if (!card.ClientRectangle.Contains(card.PointToClient(Cursor.Position)))
                        SetHover(false);
                };
                c.Cursor = Cursors.Hand;
                c.Click += click;
            }
            card.Click += click;

            return card;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  버튼 핸들러
        // ─────────────────────────────────────────────────────────────────────

        private void BtnCreateMne_Click(object sender, EventArgs e)
        {
            if (!File.Exists(TemplatePath))
            { ShowError($"템플릿 파일을 찾을 수 없습니다.\n{TemplatePath}"); return; }

            using var dlg = new SaveFileDialog
            {
                Filter   = "Excel 파일 (*.xlsx)|*.xlsx",
                Title    = "MNE 파일 저장",
                FileName = $"MNE_{DateTime.Now:yyyyMMdd}.xlsx",
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                File.Copy(TemplatePath, dlg.FileName, overwrite: true);
                MessageBox.Show($"MNE 파일이 생성되었습니다.\n{dlg.FileName}",
                    "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { ShowError($"파일 생성 오류:\n{ex.Message}"); }
        }


        private void BtnConvert_Click(object sender, EventArgs e)
        {
            using var terms = new TermsDialog();
            if (terms.ShowDialog(this) != DialogResult.OK) return;

            // main 파일 선택 (단일 파일)
            using var openDlg = new OpenFileDialog
            {
                Filter = "Excel 파일 (*.xlsx)|*.xlsx",
                Title  = "변환할 main_template_newest.xlsx 파일을 선택하세요.",
            };
            if (openDlg.ShowDialog() != DialogResult.OK) return;
            var mainFilePath = openDlg.FileName;

            using var saveDlg = new SaveFileDialog
            {
                Filter           = "XML 파일 (*.xml)|*.xml",
                Title            = "XML 파일 저장",
                FileName         = "GLOBE_OECD.xml",
                InitialDirectory = Path.GetDirectoryName(mainFilePath),
            };
            if (saveDlg.ShowDialog() != DialogResult.OK) return;

            using var logger = new Services.AppLogger();
            try
            {
                logger.WriteLine($"입력: {mainFilePath}");
                logger.WriteLine($"출력: {saveDlg.FileName}");

                var globe = new Globe.GlobeOecd
                {
                    Version     = "2.0",
                    MessageSpec = new Globe.MessageSpecType(),
                    GlobeBody   = new Globe.GlobeBodyType(),
                };

                var orchestrator  = new MappingOrchestrator();
                var mappingErrors = orchestrator.MapWorkbook(mainFilePath, globe);
                logger.WriteLine($"매핑 완료 — 오류 {mappingErrors.Count}건");
                if (mappingErrors.Count > 0)
                    logger.WriteLines(mappingErrors, "매핑 오류:");

                var xml = XmlExportService.Serialize(globe);
                File.WriteAllText(saveDlg.FileName, xml, System.Text.Encoding.UTF8);
                logger.WriteLine("XML 직렬화 완료");

                var validationErrors = ValidationUtil.Validate(globe);
                logger.WriteLine($"검증 완료 — 오류 {validationErrors.Count}건");
                if (validationErrors.Count > 0)
                    logger.WriteLines(validationErrors, "검증 오류:");

                var errorsPath = Path.ChangeExtension(saveDlg.FileName, ".errors.txt");

                if (mappingErrors.Count > 0 || validationErrors.Count > 0)
                {
                    File.WriteAllText(errorsPath,
                        $"[오류 목록] {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}"
                        + $"매핑 오류 {mappingErrors.Count}건 / 검증 오류 {validationErrors.Count}건{Environment.NewLine}{Environment.NewLine}"
                        + string.Join(Environment.NewLine, mappingErrors)
                        + Environment.NewLine
                        + string.Join(Environment.NewLine, validationErrors),
                        System.Text.Encoding.UTF8);
                    MessageBox.Show(
                        $"XML 생성 완료.\n매핑 오류: {mappingErrors.Count}건\n검증 오류: {validationErrors.Count}건\n\n오류 목록: {errorsPath}",
                        "완료 (오류 있음)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    if (File.Exists(errorsPath)) File.Delete(errorsPath);
                    MessageBox.Show("XML 생성이 완료되었습니다.", "완료",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                logger.WriteException(ex, "변환 중 예외 발생");
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("XML 변환 오류:");
                var ex2 = ex;
                while (ex2 != null)
                {
                    sb.AppendLine(ex2.Message);
                    ex2 = ex2.InnerException;
                    if (ex2 != null) sb.AppendLine("→ 원인:");
                }
                sb.AppendLine($"\n로그: {logger.LogPath}");
                ShowError(sb.ToString());
            }
        }

        private static void ShowError(string msg) =>
            MessageBox.Show(msg, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
