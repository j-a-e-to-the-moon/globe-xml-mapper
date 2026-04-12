using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace GlobeMapper
{
    public class TermsDialog : Form
    {
        private CheckBox chkAgree;
        private Button btnNext;
        private Button btnCancel;

        private static readonly Color BG     = Color.FromArgb(30, 30, 32);
        private static readonly Color BG2    = Color.FromArgb(22, 22, 24);
        private static readonly Color BG3    = Color.FromArgb(44, 44, 50);
        private static readonly Color FG     = Color.FromArgb(215, 215, 220);
        private static readonly Color BORDER = Color.FromArgb(55, 55, 62);
        private static readonly Color ACCENT = Color.FromArgb(210, 160, 0);

        public TermsDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text            = "이용 약관 동의";
            AutoScaleMode   = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            StartPosition   = FormStartPosition.CenterParent;
            ClientSize      = new Size(560, 530);
            BackColor       = BG;
            ForeColor       = FG;
            Font            = new Font("Segoe UI", 11f);

            // ── 약관 텍스트 ──────────────────────────────────────────────
            var rtb = new RichTextBox
            {
                ReadOnly    = true,
                BorderStyle = BorderStyle.None,
                ScrollBars  = RichTextBoxScrollBars.Vertical,
                BackColor   = BG3,
                ForeColor   = FG,
                Font        = new Font("Segoe UI", 11f),
                WordWrap    = true,
                Location    = new Point(16, 16),
                Size        = new Size(528, 400),
                Anchor      = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };

            var termsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "terms.txt");
            rtb.Text = File.Exists(termsPath)
                ? File.ReadAllText(termsPath)
                : "(약관 파일을 찾을 수 없습니다)";

            // ── 구분선 ────────────────────────────────────────────────────
            var divider = new Panel
            {
                BackColor = BORDER,
                Location  = new Point(16, 424),
                Size      = new Size(528, 1),
                Anchor    = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };

            // ── 동의 체크박스 ─────────────────────────────────────────────
            chkAgree = new CheckBox
            {
                Text      = "위 약관에 동의합니다.",
                AutoSize  = true,
                Location  = new Point(16, 438),
                ForeColor = FG,
                Font      = new Font("Segoe UI", 11f),
                Anchor    = AnchorStyles.Bottom | AnchorStyles.Left,
            };
            chkAgree.CheckedChanged += (s, e) => btnNext.Visible = chkAgree.Checked;

            // ── 버튼 ──────────────────────────────────────────────────────
            btnCancel = MakeBtn("취소", DialogResult.Cancel, Color.FromArgb(60, 60, 66));
            btnCancel.Location = new Point(448, 482);
            btnCancel.Anchor   = AnchorStyles.Bottom | AnchorStyles.Right;

            btnNext = MakeBtn("다음", DialogResult.OK, Color.FromArgb(55, 100, 170));
            btnNext.ForeColor = Color.White;
            btnNext.Location  = new Point(340, 482);
            btnNext.Visible   = false;
            btnNext.Anchor    = AnchorStyles.Bottom | AnchorStyles.Right;

            AcceptButton = btnNext;
            CancelButton = btnCancel;

            Controls.AddRange(new Control[] { rtb, divider, chkAgree, btnNext, btnCancel });
        }

        private static Button MakeBtn(string text, DialogResult dr, Color bg)
        {
            var hover = Color.FromArgb(
                Math.Min(bg.R + 20, 255),
                Math.Min(bg.G + 20, 255),
                Math.Min(bg.B + 20, 255));
            var btn = new Button
            {
                Text         = text,
                DialogResult = dr,
                Size         = new Size(96, 36),
                FlatStyle    = FlatStyle.Flat,
                BackColor    = bg,
                ForeColor    = Color.White,
                Font         = new Font("Segoe UI", 10.5f),
                Cursor       = Cursors.Hand,
            };
            btn.FlatAppearance.BorderSize            = 0;
            btn.FlatAppearance.MouseOverBackColor    = hover;
            return btn;
        }
    }
}
