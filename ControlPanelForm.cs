using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using GlobeMapper.Services;

namespace GlobeMapper
{
    public class ControlPanelForm : Form
    {
        private readonly ExcelController _excel;
        private Label _lblCurrentSheet;
        private Panel _bodyPanel;
        private Panel _dynamicPanel;
        private Button btnToggle;
        private Timer _sheetTracker;
        private string _lastSheetName;
        private Point _dragOffset;
        private bool _dragging;
        private bool _collapsed;

        // 시트2(1.3.1)의 최종모기업 행 블록: 3~11행
        private const int UPE_BLOCK_START = 3;
        private const int UPE_BLOCK_END = 11;
        private const int UPE_BLOCK_GAP = 2;

        // 시트5(1.3.2.2)의 제외기업 행 블록: 3~6행
        private const int EX_BLOCK_START = 3;
        private const int EX_BLOCK_END = 6;
        private const int EX_BLOCK_GAP = 2;

        // 첨부 시트 이름
        private const string ATTACH_SHEET_NAME = "1.3.2.1 첨부";

        // 별첨 번호 상태 (별첨 시트 활성 시)
        private int _selectedAttachNum = 1;

        public ControlPanelForm(ExcelController excel)
        {
            _excel = excel;
            InitializeComponent();
            StartSheetTracker();
        }

        private void InitializeComponent()
        {
            Text = "Globe XML Mapper";
            FormBorderStyle = FormBorderStyle.None;
            TopMost = true;
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = Color.FromArgb(45, 45, 48);
            ForeColor = Color.White;
            StartPosition = FormStartPosition.Manual;
            Location = new Point(Screen.PrimaryScreen.WorkingArea.Right - 380, 80);
            Size = new Size(360, 200);

            // 타이틀 바
            var titleBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 28,
                BackColor = Color.FromArgb(30, 30, 30),
                Cursor = Cursors.SizeAll
            };
            titleBar.MouseDown += (s, e) => { _dragging = true; _dragOffset = e.Location; };
            titleBar.MouseMove += (s, e) => { if (_dragging) Location = new Point(Location.X + e.X - _dragOffset.X, Location.Y + e.Y - _dragOffset.Y); };
            titleBar.MouseUp += (s, e) => _dragging = false;

            var lblTitle = new Label
            {
                Text = "Globe XML Mapper",
                AutoSize = true, Location = new Point(8, 5),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                BackColor = Color.Transparent
            };
            lblTitle.MouseDown += (s, e) => { _dragging = true; _dragOffset = new Point(e.X + lblTitle.Left, e.Y + lblTitle.Top); };
            lblTitle.MouseMove += (s, e) => { if (_dragging) Location = new Point(Location.X + e.X + lblTitle.Left - _dragOffset.X, Location.Y + e.Y + lblTitle.Top - _dragOffset.Y); };
            lblTitle.MouseUp += (s, e) => _dragging = false;

            btnToggle = new Button
            {
                Text = "─", Size = new Size(28, 22), Location = new Point(325, 3),
                FlatStyle = FlatStyle.Flat, ForeColor = Color.White,
                BackColor = Color.FromArgb(60, 60, 60), Font = new Font("Segoe UI", 8)
            };
            btnToggle.FlatAppearance.BorderSize = 0;
            btnToggle.Click += (s, e) =>
            {
                _collapsed = !_collapsed;
                _bodyPanel.Visible = !_collapsed;
                btnToggle.Text = _collapsed ? "□" : "─";
                ResizeToFit();
            };

            titleBar.Controls.Add(lblTitle);
            titleBar.Controls.Add(btnToggle);

            // 본문 패널
            _bodyPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8) };

            var y = 4;

            // 현재 시트 표시
            _lblCurrentSheet = new Label
            {
                Text = "현재 시트: -",
                Location = new Point(8, y), Size = new Size(336, 20),
                ForeColor = Color.FromArgb(100, 200, 255),
                Font = new Font("Segoe UI", 8.5f)
            };
            _bodyPanel.Controls.Add(_lblCurrentSheet);
            y += 24;

            // 구분선
            _bodyPanel.Controls.Add(new Panel
            {
                Location = new Point(8, y), Size = new Size(336, 1),
                BackColor = Color.FromArgb(70, 70, 70)
            });
            y += 6;

            // 동적 버튼 영역 (시트에 따라 내용 변경)
            _dynamicPanel = new Panel
            {
                Location = new Point(0, y), Size = new Size(360, 100),
                BackColor = Color.Transparent
            };
            _bodyPanel.Controls.Add(_dynamicPanel);

            Controls.Add(_bodyPanel);
            Controls.Add(titleBar);
        }

        #region 시트별 동적 UI

        private void UpdateDynamicPanel(string sheetName)
        {
            _dynamicPanel.Controls.Clear();
            var y = 0;

            if (sheetName == "1.3.1")
            {
                // 시트2: 최종모기업 추가/삭제 + 시트 초기화
                var count = _excel.GetRowBlockCount(sheetName);
                y = AddSectionLabel("최종모기업", $"{count}개", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddRowBlock(sheetName, UPE_BLOCK_START, UPE_BLOCK_END, UPE_BLOCK_GAP); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (count <= 1) { MessageBox.Show("최소 1개는 유지해야 합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        if (MessageBox.Show("마지막 최종모기업을 삭제하시겠습니까?", "확인", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                        _excel.RemoveRowBlock(sheetName, UPE_BLOCK_START, UPE_BLOCK_END, UPE_BLOCK_GAP);
                        UpdateDynamicPanel(sheetName);
                    })
                );
                y += 4;
                y = AddActionButton("시트 초기화", Color.FromArgb(100, 100, 100), y, () =>
                {
                    if (MessageBox.Show("시트를 초기 상태로 되돌리시겠습니까?", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
                    _excel.ResetSheet(sheetName, UPE_BLOCK_START, UPE_BLOCK_END, UPE_BLOCK_GAP);
                    UpdateDynamicPanel(sheetName);
                });
            }
            else if (sheetName != null && sheetName.Contains("첨부"))
            {
                // 첨부 시트 (1.3.2.1 첨부)
                var ceCount = _excel.GetCeBlockCount("1.3.2.1");
                if (_selectedAttachNum > ceCount) _selectedAttachNum = 1;
                if (_selectedAttachNum < 1) _selectedAttachNum = 1;

                y = AddAttachSelector(y, ceCount, sheetName);
                y += 4;

                var ownerCount = _excel.GetOwnerRowCount(sheetName, _selectedAttachNum);
                y = AddSectionLabel($"첨부{_selectedAttachNum} 주주", $"{ownerCount}행", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddOwnerRow(sheetName, _selectedAttachNum); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (ownerCount <= 0) { MessageBox.Show("삭제할 행이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        _excel.RemoveOwnerRow(sheetName, _selectedAttachNum);
                        UpdateDynamicPanel(sheetName);
                    })
                );
            }
            else if (sheetName == "1.3.2.1")
            {
                // 시트3: 구성기업 추가/삭제 + 시트 초기화
                var count = _excel.GetCeBlockCount(sheetName);
                y = AddSectionLabel("구성기업", $"{count}개", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddCeBlock(sheetName, ATTACH_SHEET_NAME); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (count <= 1) { MessageBox.Show("최소 1개는 유지해야 합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        if (MessageBox.Show("마지막 구성기업을 삭제하시겠습니까?", "확인", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                        _excel.RemoveCeBlock(sheetName, ATTACH_SHEET_NAME);
                        UpdateDynamicPanel(sheetName);
                    })
                );
                y += 4;
                y = AddActionButton("시트 초기화", Color.FromArgb(100, 100, 100), y, () =>
                {
                    if (MessageBox.Show("시트를 초기 상태로 되돌리시겠습니까?\n모든 구성기업 및 첨부 데이터가 삭제됩니다.", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
                    _excel.ResetCeSheet(sheetName, ATTACH_SHEET_NAME);
                    UpdateDynamicPanel(sheetName);
                });
            }
            else if (sheetName == "1.3.2.2")
            {
                // 시트5: 제외기업 추가/삭제 + 시트 초기화
                var count = _excel.GetRowBlockCount(sheetName);
                y = AddSectionLabel("제외기업", $"{count}개", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddRowBlock(sheetName, EX_BLOCK_START, EX_BLOCK_END, EX_BLOCK_GAP); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (count <= 1) { MessageBox.Show("최소 1개는 유지해야 합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        if (MessageBox.Show("마지막 제외기업을 삭제하시겠습니까?", "확인", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                        _excel.RemoveRowBlock(sheetName, EX_BLOCK_START, EX_BLOCK_END, EX_BLOCK_GAP);
                        UpdateDynamicPanel(sheetName);
                    })
                );
                y += 4;
                y = AddActionButton("시트 초기화", Color.FromArgb(100, 100, 100), y, () =>
                {
                    if (MessageBox.Show("시트를 초기 상태로 되돌리시겠습니까?", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
                    _excel.ResetSheet(sheetName, EX_BLOCK_START, EX_BLOCK_END, EX_BLOCK_GAP);
                    UpdateDynamicPanel(sheetName);
                });
            }

            else if (sheetName == "2")
            {
                // 시트 2: 국가별 적용면제 페이지 추가/삭제
                var count = _excel.GetRowBlockCount(sheetName);
                y = AddSectionLabel("국가별 적용면제", $"{count}개", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddSheet2Block(sheetName); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (count <= 1) { MessageBox.Show("최소 1개는 유지해야 합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        if (MessageBox.Show("마지막 페이지를 삭제하시겠습니까?", "확인", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                        _excel.RemoveSheet2Block(sheetName);
                        UpdateDynamicPanel(sheetName);
                    })
                );
                y += 4;
                y = AddActionButton("시트 초기화", Color.FromArgb(100, 100, 100), y, () =>
                {
                    if (MessageBox.Show("시트를 초기 상태로 되돌리시겠습니까?", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
                    _excel.ResetSheet2(sheetName);
                    UpdateDynamicPanel(sheetName);
                });
            }
            else if (sheetName == "1.4")
            {
                // 1.4: 정보 요약 행 추가/삭제 (헤더 3행, 데이터 시작 4행)
                var count = _excel.GetRowBlockCount(sheetName);
                y = AddSectionLabel("정보 요약", $"{count}행", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddSimpleRowByMeta(sheetName, 4); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (count <= 1) { MessageBox.Show("최소 1행은 유지해야 합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        _excel.RemoveSimpleRowByMeta(sheetName, 4);
                        UpdateDynamicPanel(sheetName);
                    })
                );
            }
            else if (sheetName == "1.3.3")
            {
                // 1.3.3: 기업구조 변동 행 추가/삭제 (헤더 6행, 데이터 시작 7행)
                var count = _excel.GetRowBlockCount(sheetName);
                y = AddSectionLabel("기업구조 변동", $"{count}행", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddSimpleRowByMeta(sheetName, 7); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (count <= 1) { MessageBox.Show("최소 1행은 유지해야 합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        _excel.RemoveSimpleRowByMeta(sheetName, 7);
                        UpdateDynamicPanel(sheetName);
                    })
                );
            }

            else if (sheetName == "3.1~3.2.3.2")
            {
                // 시트3: 페이지 추가 + 3개 행 추가 영역
                var pageCount = _excel.GetRowBlockCount(sheetName);
                y = AddSectionLabel("글로벌최저한세 계산", $"{pageCount}페이지", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddSheet3Page(sheetName); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (pageCount <= 1) { MessageBox.Show("최소 1페이지는 유지해야 합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        if (MessageBox.Show("마지막 페이지를 삭제하시겠습니까?", "확인", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                        _excel.RemoveSheet3Page(sheetName);
                        UpdateDynamicPanel(sheetName);
                    })
                );

                y += 6;

                // (c) 통합형피지배 (101행)
                var cfcCount = _excel.GetSheet3RowCount(sheetName, "cfc");
                y = AddSectionLabel("통합형피지배 행", $"{cfcCount}행", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddSheet3Row(sheetName, "cfc", 101); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (cfcCount <= 1) { MessageBox.Show("최소 1행은 유지해야 합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        _excel.RemoveSheet3Row(sheetName, "cfc", 101);
                        UpdateDynamicPanel(sheetName);
                    })
                );

                // (c) 결손금 소급공제 (144행)
                var cbCount = _excel.GetSheet3RowCount(sheetName, "carryback");
                y = AddSectionLabel("결손금 소급공제 행", $"{cbCount}행", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddSheet3Row(sheetName, "carryback", 144); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (cbCount <= 1) { MessageBox.Show("최소 1행은 유지해야 합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        _excel.RemoveSheet3Row(sheetName, "carryback", 144);
                        UpdateDynamicPanel(sheetName);
                    })
                );

                // (b) 제89조 (176행)
                var artCount = _excel.GetSheet3RowCount(sheetName, "art89");
                y = AddSectionLabel("제89조 행", $"{artCount}행", y);
                y = AddButtonRow(y,
                    ("+", Color.LimeGreen, () => { _excel.AddSheet3Row(sheetName, "art89", 176); UpdateDynamicPanel(sheetName); }),
                    ("−", Color.Tomato, () =>
                    {
                        if (artCount <= 1) { MessageBox.Show("최소 1행은 유지해야 합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                        _excel.RemoveSheet3Row(sheetName, "art89", 176);
                        UpdateDynamicPanel(sheetName);
                    })
                );

                y += 4;
                y = AddActionButton("시트 초기화", Color.FromArgb(100, 100, 100), y, () =>
                {
                    if (MessageBox.Show("시트를 초기 상태로 되돌리시겠습니까?", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
                    _excel.ResetSheet3(sheetName);
                    UpdateDynamicPanel(sheetName);
                });
            }

            // 공통: 엑셀 종료
            y += 10;
            y = AddActionButton("엑셀 종료하기", Color.FromArgb(80, 80, 80), y, BtnCloseExcel_Click);

            _dynamicPanel.Height = y + 4;
            ResizeToFit();
        }

        private bool IsSheetSection(string section, string sheetName)
        {
            var sheets = _excel.GetSectionSheets(section);
            return sheets.Contains(sheetName) || sheetName.Contains($"({(section == "1.3.2.1" ? "2" : "3")})");
        }

        #endregion

        #region UI 헬퍼

        private int AddSectionLabel(string name, string countText, int y)
        {
            _dynamicPanel.Controls.Add(new Label
            {
                Text = name, Location = new Point(8, y + 3), AutoSize = true,
                ForeColor = Color.LightGray, Font = new Font("Segoe UI", 8.5f)
            });
            _dynamicPanel.Controls.Add(new Label
            {
                Text = countText, Location = new Point(200, y + 3), AutoSize = true,
                ForeColor = Color.White, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
            });
            return y + 24;
        }

        private int AddAttachSelector(int y, int maxNum, string sheetName)
        {
            var lbl = new Label
            {
                Text = "별첨 번호:",
                Location = new Point(8, y + 4), AutoSize = true,
                ForeColor = Color.LightGray, Font = new Font("Segoe UI", 8.5f)
            };
            _dynamicPanel.Controls.Add(lbl);

            var btnPrev = MakeSmallButton("◀", Color.White, new Point(90, y), () =>
            {
                if (_selectedAttachNum > 1) { _selectedAttachNum--; UpdateDynamicPanel(sheetName); }
            });

            var txtNum = new TextBox
            {
                Text = _selectedAttachNum.ToString(),
                Location = new Point(120, y), Size = new Size(35, 24),
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9)
            };
            txtNum.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (int.TryParse(txtNum.Text, out var n) && n >= 1 && n <= maxNum)
                    { _selectedAttachNum = n; UpdateDynamicPanel(sheetName); }
                    else
                    { txtNum.Text = _selectedAttachNum.ToString(); }
                    e.SuppressKeyPress = true;
                }
            };

            var btnNext = MakeSmallButton("▶", Color.White, new Point(160, y), () =>
            {
                if (_selectedAttachNum < maxNum) { _selectedAttachNum++; UpdateDynamicPanel(sheetName); }
            });

            var lblMax = new Label
            {
                Text = $"/ {maxNum}",
                Location = new Point(190, y + 4), AutoSize = true,
                ForeColor = Color.Gray, Font = new Font("Segoe UI", 8f)
            };

            _dynamicPanel.Controls.AddRange(new Control[] { btnPrev, txtNum, btnNext, lblMax });
            return y + 30;
        }

        private int AddButtonRow(int y, (string text, Color color, Action click) btn1, (string text, Color color, Action click) btn2)
        {
            var b1 = MakeSmallButton(btn1.text, btn1.color, new Point(290, y - 22), btn1.click);
            var b2 = MakeSmallButton(btn2.text, btn2.color, new Point(320, y - 22), btn2.click);
            _dynamicPanel.Controls.Add(b1);
            _dynamicPanel.Controls.Add(b2);
            return y;
        }

        private Button MakeSmallButton(string text, Color foreColor, Point location, Action click)
        {
            var btn = new Button
            {
                Text = text, Size = new Size(26, 24), Location = location,
                FlatStyle = FlatStyle.Flat, ForeColor = foreColor,
                BackColor = Color.FromArgb(60, 60, 60),
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, e) => { try { click(); } catch (Exception ex) { MessageBox.Show(ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error); } };
            return btn;
        }

        private int AddActionButton(string text, Color bgColor, int y, Action click)
        {
            var btn = new Button
            {
                Text = text, Location = new Point(8, y), Size = new Size(336, 30),
                FlatStyle = FlatStyle.Flat, BackColor = bgColor,
                ForeColor = Color.White, Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, e) => { try { click(); } catch (Exception ex) { MessageBox.Show(ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error); } };
            _dynamicPanel.Controls.Add(btn);
            return y + 34;
        }

        private void ResizeToFit()
        {
            if (_collapsed)
                Height = 28;
            else
                Height = 28 + 34 + _dynamicPanel.Height + 16; // titleBar + currentSheet+sep + dynamic + padding
        }

        #endregion

        #region 시트 트래커 (1초 폴링)

        private void StartSheetTracker()
        {
            _sheetTracker = new Timer { Interval = 1000 };
            _sheetTracker.Tick += (s, e) =>
            {
                try
                {
                    if (!_excel.IsOpen)
                    {
                        _sheetTracker.Stop();
                        Close();
                        return;
                    }

                    var current = _excel.GetActiveSheetName();

                    // _META 시트로 이동 방지
                    if (current == ExcelController.MetaSheetName)
                    {
                        _excel.ActivateSheet(_lastSheetName != null ? (object)_lastSheetName : 1);
                        return;
                    }

                    if (current != null && current != _lastSheetName)
                    {
                        _lastSheetName = current;
                        _lblCurrentSheet.Text = $"현재 시트: {current}";
                        UpdateDynamicPanel(current);
                    }
                }
                catch
                {
                    _sheetTracker.Stop();
                    Close();
                }
            };
            _sheetTracker.Start();

            // 초기 한 번 실행
            var initial = _excel.GetActiveSheetName();
            if (initial != null)
            {
                _lastSheetName = initial;
                _lblCurrentSheet.Text = $"현재 시트: {initial}";
                UpdateDynamicPanel(initial);
            }
        }

        #endregion

        #region 액션

        private void BtnConvert_Click()
        {
            using var terms = new TermsDialog();
            if (terms.ShowDialog(this) != DialogResult.OK) return;

            using var dlg = new SaveFileDialog
            {
                Filter = "XML 파일 (*.xml)|*.xml",
                Title = "XML 파일 저장",
                FileName = "GLOBE_OECD.xml"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            var filePath = _excel.GetFilePathForMapping();

            var globe = new Globe.GlobeOecd
            {
                Version = "2.0",
                MessageSpec = new Globe.MessageSpecType(),
                GlobeBody = new Globe.GlobeBodyType()
            };

            var orchestrator = new MappingOrchestrator();
            var mappingErrors = orchestrator.MapWorkbook(filePath, globe);

            var xml = XmlExportService.Serialize(globe);
            File.WriteAllText(dlg.FileName, xml, System.Text.Encoding.UTF8);

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

            var errorsPath = Path.ChangeExtension(dlg.FileName, ".errors.txt");
            if (allErrors.Count > 0)
            {
                File.WriteAllText(errorsPath,
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
                if (File.Exists(errorsPath)) File.Delete(errorsPath);
                MessageBox.Show("XML 생성이 완료되었습니다.", "완료",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnCloseExcel_Click()
        {
            _sheetTracker?.Stop();
            _excel.CloseWithSavePrompt();
            Close();
        }

        #endregion

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _sheetTracker?.Stop();
            _sheetTracker?.Dispose();
            base.OnFormClosing(e);
        }
    }
}
