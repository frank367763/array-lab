using System.Globalization;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ArrayLab
{
    public partial class Form1 : Form
    {
        private double[] initialArray;   // Изначальный (неприкасаемый)
        private double[] editableArray;  // Рабочий (редактируемый)
        private double[] resultArray;    // Итоговый (после обработки)

        private string inputMethod = "Отсутствует";
        private bool isManuallyEdited = false;
        private bool isUpdatingGrid = false;
        private static readonly Random globalRandom = new Random();
        private DataGridView dgvInitial;
        private DataGridView dgvEditable;
        private DataGridView dgvResult;

        private TextBox txtStatus;
        private Button btnCalculate;
        private Label lblArrayInfo;
        private Label lblInitial, lblEditable, lblResult;
        private static readonly Color CLR_BG = Color.FromArgb(245, 246, 250);
        private static readonly Color CLR_PANEL = Color.White;
        private static readonly Color CLR_ACCENT = Color.FromArgb(0, 95, 184);
        private static readonly Color CLR_ACCENT_H = Color.FromArgb(0, 120, 215);
        private static readonly Color CLR_SUCCESS = Color.FromArgb(16, 124, 16);
        private static readonly Color CLR_SUCCESS_H = Color.FromArgb(22, 160, 22);
        private static readonly Color CLR_SECONDARY = Color.FromArgb(85, 95, 109);
        private static readonly Color CLR_SECONDARY_H = Color.FromArgb(110, 120, 135);
        private static readonly Color CLR_BORDER = Color.FromArgb(218, 220, 228);
        private static readonly Color CLR_HEADER_TXT = Color.FromArgb(43, 56, 78);
        private static readonly Color CLR_HINT_TXT = Color.FromArgb(120, 130, 148);
        private static readonly Color CLR_TERMINAL = Color.FromArgb(20, 22, 30);
        private static readonly Color CLR_INFO_OK = Color.FromArgb(16, 124, 16);
        private static readonly Color CLR_INFO_EDITED = Color.FromArgb(196, 43, 28);
        private static readonly Color CLR_READONLY_BG = Color.FromArgb(240, 242, 245);
        private static readonly Color CLR_RESULT_BG = Color.FromArgb(235, 248, 240);
        public Form1()
        {
            this.Opacity = 0;
            SetupModernInterface();
            System.Windows.Forms.Timer ft = new System.Windows.Forms.Timer { Interval = 15 };
            ft.Tick += (s, e) => { if (this.Opacity >= 1.0) ft.Stop(); else this.Opacity += 0.07; };
            this.Load += (s, e) => ft.Start();
        }
        private void SetupModernInterface()
        {
            this.Text = "ArrayLab — Программа обработки массива (Вариант 9)";
            this.Size = new Size(940, 880);
            this.MinimumSize = new Size(940, 880);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = CLR_BG;
            this.Font = new Font("Segoe UI", 9);
            this.DoubleBuffered = true;

            TabControl tabs = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10) };
            TabPage tab1 = new TabPage("  ▶  Работа с массивом  ") { UseVisualStyleBackColor = true };
            TabPage tab2 = new TabPage("  ?  Справочник  ") { UseVisualStyleBackColor = true };
            tabs.TabPages.AddRange(new TabPage[] { tab1, tab2 });
            this.Controls.Add(tabs);

            Panel inner = new Panel { Dock = DockStyle.Fill, BackColor = CLR_BG, AutoScroll = true };
            tab1.Controls.Add(inner);
            BuildMainTab(inner);
            BuildHelpTab(tab2);
        }

        private void BuildMainTab(Panel c)
        {
            int px = 10, pw = 900, y = 8;
            Panel ph = MakePanel(px, y, pw, 58, CLR_ACCENT);
            ph.Controls.Add(MakeLbl("Программа обработки массива", Color.White,
                new Font("Segoe UI", 14, FontStyle.Bold), 15, 7, autoSize: true));
            ph.Controls.Add(MakeLbl(
                "Вариант 9: находит Min и Max, обнуляет элементы между ними. N ∈ [1;20], значения ∈ [-1000;1000].",
                Color.FromArgb(200, 220, 255), new Font("Segoe UI", 8), 15, 35, autoSize: true));
            c.Controls.Add(ph); y += 66;
            Panel pi = MakePanel(px, y, pw, 125);
            AddPanelHeader(pi, "1  Выберите способ ввода данных");

            Button btnManual = MakeBtn("Ручной ввод", 15, 42, CLR_ACCENT, CLR_ACCENT_H);
            Button btnRandom = MakeBtn("Генерация", 198, 42, CLR_ACCENT, CLR_ACCENT_H);
            Button btnInTxt = MakeBtn("Загрузить .txt", 381, 42, CLR_SECONDARY, CLR_SECONDARY_H);
            Button btnInXls = MakeBtn("Импорт Excel", 564, 42, CLR_SECONDARY, CLR_SECONDARY_H);

            btnManual.Click += (s, e) => PrepareManualInput();
            btnRandom.Click += (s, e) => GenerateRandom();
            btnInTxt.Click += (s, e) => LoadTxt();
            btnInXls.Click += (s, e) => LoadExcel();

            pi.Controls.AddRange(new Control[] {
                btnManual, btnRandom, btnInTxt, btnInXls,
                MakeHint("Откроет окно ввода, где программа сама определит размер по факту ввода элементов.", 15, 78),
                MakeHint("Случайные числа. Можно выбрать случайный размер N.", 198, 78),
                MakeHint("Строка 1: N. Строка 2: числа через пробел.", 381, 78),
                MakeHint("A1 = N, затем элементы в столбик (A2..A(N+1)).", 564, 78)
            });
            c.Controls.Add(pi); y += 133;
            Panel pa = MakePanel(px, y, pw, 380);
            AddPanelHeader(pa, "2  Массивы");

            lblArrayInfo = new Label
            {
                Left = 15,
                Top = 35,
                Width = 870,
                Height = 18,
                AutoSize = false,
                Font = new Font("Segoe UI", 8, FontStyle.Bold | FontStyle.Italic),
                ForeColor = CLR_INFO_OK,
                Text = "Способ ввода: Отсутствует"
            };
            pa.Controls.Add(lblArrayInfo);

            int dgvY = 55;
            lblInitial = MakeLbl("Исходный массив (только чтение):", CLR_HEADER_TXT, new Font("Segoe UI", 9, FontStyle.Bold), 15, dgvY);
            dgvInitial = CreateGrid(15, dgvY + 20, true, CLR_READONLY_BG);
            pa.Controls.AddRange(new Control[] { lblInitial, dgvInitial }); dgvY += 80;
            lblEditable = MakeLbl("Рабочий массив (можно редактировать):", CLR_HEADER_TXT, new Font("Segoe UI", 9, FontStyle.Bold), 15, dgvY);
            dgvEditable = CreateGrid(15, dgvY + 20, false, Color.White);
            dgvEditable.CellValueChanged += (s, e) => { if (!isUpdatingGrid) { isManuallyEdited = true; UpdateArrayInfo(); } };
            dgvEditable.CellValidating += DGV_CellValidating;
            dgvEditable.CellEndEdit += (s, e) => { if (e.RowIndex == 0) dgvEditable.Rows[0].ErrorText = ""; };
            dgvEditable.DataError += (s, e) => { e.ThrowException = false; };
            pa.Controls.AddRange(new Control[] { lblEditable, dgvEditable }); dgvY += 85;
            btnCalculate = new Button
            {
                Text = "▶  ОБНУЛИТЬ ЭЛЕМЕНТЫ МЕЖДУ MIN И MAX",
                Top = dgvY,
                Left = 15,
                Width = 870,
                Height = 35,
                BackColor = CLR_SUCCESS,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Enabled = false
            };
            btnCalculate.FlatAppearance.BorderSize = 0;
            btnCalculate.FlatAppearance.MouseOverBackColor = CLR_SUCCESS_H;
            btnCalculate.Click += (s, e) => { AnimateButton(btnCalculate, CLR_SUCCESS); ProcessArray(); };
            pa.Controls.Add(btnCalculate); dgvY += 45;
            lblResult = MakeLbl("Итоговый результат (для сохранения):", CLR_SUCCESS, new Font("Segoe UI", 9, FontStyle.Bold), 15, dgvY);
            dgvResult = CreateGrid(15, dgvY + 20, true, CLR_RESULT_BG);
            pa.Controls.AddRange(new Control[] { lblResult, dgvResult });

            c.Controls.Add(pa); y += 388;
            Panel pe = MakePanel(px, y, pw, 82);
            AddPanelHeader(pe, "3  Сохранение результата");
            Button btnOutTxt = MakeBtn("Сохранить в .txt", 15, 36, CLR_SECONDARY, CLR_SECONDARY_H);
            Button btnOutXls = MakeBtn("Экспорт в Excel", 198, 36, CLR_SECONDARY, CLR_SECONDARY_H);
            btnOutTxt.Click += (s, e) => SaveTxt();
            btnOutXls.Click += (s, e) => SaveExcel();

            pe.Controls.AddRange(new Control[] {
                btnOutTxt, btnOutXls,
                MakeHint("Внимание: Сохраняется только «Итоговый результат»!", 381, 42)
            });
            c.Controls.Add(pe); y += 90;
            Panel pl = new Panel { Left = px, Top = y, Width = pw, Height = 135, BackColor = CLR_TERMINAL };
            pl.Controls.Add(new Label
            {
                Text = "  ●  Лог работы программы",
                ForeColor = Color.FromArgb(100, 120, 160),
                Font = new Font("Consolas", 8),
                Left = 0,
                Top = 0,
                Width = pw,
                Height = 20,
                BackColor = Color.FromArgb(12, 14, 20),
                Padding = new Padding(5, 3, 0, 0)
            });
            txtStatus = new TextBox
            {
                Left = 4,
                Top = 22,
                Width = pw - 4,
                Height = 110,
                BackColor = CLR_TERMINAL,
                ForeColor = Color.FromArgb(74, 246, 38),
                Font = new Font("Consolas", 9),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BorderStyle = BorderStyle.None
            };
            pl.Controls.Add(txtStatus);
            c.Controls.Add(pl);
        }

        private DataGridView CreateGrid(int x, int y, bool isReadonly, Color bgColor)
        {
            var g = new DataGridView
            {
                Left = x,
                Top = y,
                Width = 870,
                Height = 68,
                AllowUserToAddRows = false,
                ReadOnly = isReadonly,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
                BackgroundColor = bgColor,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor = CLR_BORDER,
                RowHeadersVisible = false,
                EnableHeadersVisualStyles = false,
                SelectionMode = isReadonly ? DataGridViewSelectionMode.FullRowSelect : DataGridViewSelectionMode.CellSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ScrollBars = ScrollBars.Horizontal
            };
            ApplyModernGridStyle(g);
            return g;
        }
        private void BuildHelpTab(TabPage tab)
        {
            Panel wrapper = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(0) };
            tab.Controls.Add(wrapper);

            RichTextBox rtb = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Font = new Font("Segoe UI", 10),
                Padding = new Padding(25)
            };
            wrapper.Controls.Add(rtb);

            rtb.SelectionFont = new Font("Segoe UI", 18, FontStyle.Bold);
            rtb.SelectionColor = CLR_ACCENT;
            rtb.AppendText("ArrayLab — Полное руководство пользователя\n");
            rtb.SelectionFont = new Font("Segoe UI", 9);
            rtb.SelectionColor = CLR_HINT_TXT;
            rtb.AppendText("Программный комплекс обработки одномерных массивов (Вариант 9)\n\n");

            HelpSection(rtb, "1. Что делает программа?",
                "Основная задача — автоматизация обработки числовых последовательностей.\n" +
                "Алгоритм находит в массиве МИНИМАЛЬНЫЙ и МАКСИМАЛЬНЫЙ элементы, " +
                "после чего обнуляет (присваивает значение 0.0) все элементы, расположенные между ними.\n\n" +
                "Пример работы:\n" +
                "   Вход: [ 5.0,  1.2,  -10.0,  4.4,  20.0,  0.0 ]\n" +
                "   Min: -10.0 (индекс 2), Max: 20.0 (индекс 4)\n" +
                "   Элемент между ними: 4.4 (индекс 3)\n" +
                "   Результат: [ 5.0, 1.2, -10.0, 0.0, 20.0, 0.0 ]");

            HelpSection(rtb, "2. Концепция «Трёх таблиц»",
                "Для обеспечения надежности и возможности сравнения данных, интерфейс разделен на 3 зоны:\n\n" +
                "• ИСХОДНЫЙ МАССИВ (ReadOnly): Отображает данные в том виде, в котором они были получены из файла или генератора. " +
                "Эту таблицу нельзя редактировать. Она служит эталоном.\n\n" +
                "• РАБОЧИЙ МАССИВ (Editable): Копия исходного массива. Здесь вы можете вручную изменить любое значение, " +
                "если заметили ошибку во входных данных. Именно из этой таблицы берутся числа для расчёта.\n\n" +
                "• ИТОГОВЫЙ РЕЗУЛЬТАТ (Success Style): Появляется после нажатия кнопки «Обнулить». " +
                "Содержит финальное состояние массива после работы алгоритма. Только эти данные подлежат сохранению.");

            HelpSection(rtb, "3. Ограничения и ОДЗ",
                "Программа строго контролирует вводимые данные:\n" +
                "  • Размер массива (N): от 1 до 20 элементов.\n" +
                "  • Значение элементов: от -1000.00 до 1000.00.\n" +
                "  • Точность: до 2 знаков после запятой.\n" +
                "  • Разделитель: допускается использование точки или запятой (автоматически приводится к единому стандарту).");

            HelpSection(rtb, "4. Подробно о способах ввода",
                "1. РУЧНОЙ ВВОД:\n" +
                "   Открывает окно ввода элементов. Размер N вычисляется автоматически по факту ввода чисел через пробел.\n\n" +
                "2. ГЕНЕРАЦИЯ:\n" +
                "   Позволяет быстро создать массив. Использует защищенный статический генератор во избежание перекоса частот.\n\n" +
                "3. ЗАГРУЗКА .TXT:\n" +
                "   Файл должен иметь структуру:\n" +
                "   Строка 1: Целое число (N)\n" +
                "   Строка 2: Числа через пробел (элементы)\n\n" +
                "4. ИМПОРТ EXCEL:\n" +
                "   Ячейка A1: N. Ячейки A2, A3...: элементы массива по вертикали.");

            HelpSection(rtb, "5. Редактирование и валидация в таблице",
                "Вы можете менять значения в «Рабочем массиве»:\n" +
                "  • Если вы введете текст вместо числа или число вне диапазона, ячейка подсветится красным.\n" +
                "  • В статусной строке появится надпись «Отредактирован вручную: Да».");

            HelpSection(rtb, "6. Особые случаи алгоритма",
                "• Min и Max равны: Если все элементы одинаковы, программа предложит обнулить всё, кроме первого и последнего.\n" +
                "• Несколько Min/Max: В случае дублей программа предложит детальный выбор области обнуления.\n" +
                "• Соседние элементы: Если Min и Max стоят рядом, обнуление не производится.");

            HelpSection(rtb, "7. Справочник кодов ошибок",
                "При возникновении проблем программа выдаст код:\n" +
                "  3.2.6.1 — Некорректный размер N (вне 1..20).\n" +
                "  3.2.6.2 — Элемент выходит за границы [-1000..1000].\n" +
                "  3.2.6.3 — Ошибка структуры файла (.txt или формат).\n" +
                "  3.2.6.4 — Количество чисел не совпадает с указанным N.\n" +
                "  3.2.6.5 — Файл не найден или занят другим процессом.\n" +
                "  3.2.6.6 — Введено нечисловое значение (буквы, спецсимволы).\n" +
                "  3.2.6.7 — Ошибка при записи файла на диск.\n" +
                "  3.2.6.8 — Критическая ошибка парсинга Excel.\n" +
                "  3.2.6.9 — Обнаружена пустая ячейка в Excel.");

            rtb.SelectionStart = 0;
            rtb.ScrollToCaret();
        }

        private void HelpSection(RichTextBox rtb, string title, string body)
        {
            rtb.SelectionFont = new Font("Segoe UI", 12, FontStyle.Bold);
            rtb.SelectionColor = CLR_HEADER_TXT;
            rtb.AppendText("\n" + title + "\n");

            rtb.SelectionFont = new Font("Segoe UI", 9);
            rtb.SelectionColor = CLR_ACCENT;
            rtb.AppendText("─────────────────────────────────────────────────────────────\n");

            rtb.SelectionFont = new Font("Segoe UI", 10);
            rtb.SelectionColor = Color.FromArgb(50, 50, 50);
            rtb.AppendText(body + "\n");
        }
        private Panel MakePanel(int x, int y, int w, int h, Color? bg = null)
        {
            var p = new Panel { Left = x, Top = y, Width = w, Height = h, BackColor = bg ?? CLR_PANEL };
            if (bg == null)
                p.Paint += (s, e) => ControlPaint.DrawBorder(e.Graphics, p.ClientRectangle,
                    CLR_BORDER, 1, ButtonBorderStyle.Solid, CLR_BORDER, 1, ButtonBorderStyle.Solid,
                    CLR_BORDER, 1, ButtonBorderStyle.Solid, CLR_BORDER, 1, ButtonBorderStyle.Solid);
            return p;
        }

        private void AddPanelHeader(Panel p, string text)
        {
            p.Controls.Add(new Label
            {
                Text = text,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = CLR_HEADER_TXT,
                Left = 15,
                Top = 10,
                AutoSize = true
            });
            p.Controls.Add(new Panel { Left = 0, Top = 28, Width = p.Width, Height = 1, BackColor = CLR_BORDER });
        }

        private Button MakeBtn(string text, int x, int y, Color bg, Color hoverBg)
        {
            var b = new Button
            {
                Text = text,
                Left = x,
                Top = y,
                Width = 168,
                Height = 33,
                BackColor = bg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = hoverBg;
            return b;
        }

        private Label MakeHint(string text, int x, int y) =>
            new Label
            {
                Text = text,
                Left = x,
                Top = y,
                Width = 168,
                Height = 40,
                Font = new Font("Segoe UI", 7),
                ForeColor = CLR_HINT_TXT
            };

        private Label MakeLbl(string text, Color color, Font font, int x, int y,
            bool autoSize = false, int w = 0, int h = 0)
        {
            var l = new Label { Text = text, ForeColor = color, Font = font, Left = x, Top = y, AutoSize = autoSize };
            if (!autoSize) { if (w > 0) l.Width = w; if (h > 0) l.Height = h; }
            return l;
        }

        private void ApplyModernGridStyle(DataGridView g)
        {
            g.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = CLR_HEADER_TXT,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                Padding = new Padding(0, 3, 0, 3)
            };
            g.ColumnHeadersHeight = 25;
            g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            g.DefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = g.BackgroundColor,
                ForeColor = CLR_HEADER_TXT,
                Font = new Font("Segoe UI", 9),
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                SelectionBackColor = Color.FromArgb(220, 235, 255),
                SelectionForeColor = CLR_HEADER_TXT
            };
        }

        private void AnimateButton(Button btn, Color orig)
        {
            btn.BackColor = ControlPaint.Dark(orig, 0.15f);
            System.Windows.Forms.Timer t = new System.Windows.Forms.Timer { Interval = 100 };
            t.Tick += (s, e) => { btn.BackColor = orig; t.Stop(); };
            t.Start();
        }
        private void ClearAllData()
        {
            initialArray = null; editableArray = null; resultArray = null;
            UpdateGrid(dgvInitial, null);
            UpdateGrid(dgvEditable, null);
            UpdateGrid(dgvResult, null);
            btnCalculate.Enabled = false;
            isManuallyEdited = false;
            UpdateArrayInfo();
            Log("Таблицы очищены из-за ошибки или начала нового ввода.");
        }

        private void Log(string msg, bool isError = false)
        {
            txtStatus.AppendText($"{DateTime.Now.ToLongTimeString()}  {(isError ? "✖ ОШИБКА" : "✔ OK")}  {msg}\r\n");
            txtStatus.ScrollToCaret();
        }

        private void ShowError(string code, string desc)
        {
            Log($"{code} — {desc}", true);
            MessageBox.Show($"Код ошибки: {code}\n{desc}", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void UpdateArrayInfo()
        {
            string sizeStr = initialArray != null ? $"Размер N: {initialArray.Length}" : "Массив пуст";
            lblArrayInfo.ForeColor = isManuallyEdited ? CLR_INFO_EDITED : CLR_INFO_OK;
            lblArrayInfo.Text = $"Способ ввода: {inputMethod}  |  Отредактирован вручную: {(isManuallyEdited ? "Да" : "Нет")}  |  {sizeStr}";
        }
        private bool TryParseElements(string[] parts, int n, string srcLabel, out double[] arr)
        {
            arr = null;
            if (parts.Length != n) { ShowError("3.2.6.4", $"Элементов {parts.Length}, ожидалось {n}."); return false; }
            arr = new double[n];
            for (int i = 0; i < n; i++)
            {
                string v = parts[i].Replace(',', '.'); 
                
                if (!double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                {
                    ShowError("3.2.6.6", $"{srcLabel}[{i}]: «{parts[i]}» — не число."); arr = null; return false;
                }
                if (val < -1000 || val > 1000)
                {
                    ShowError("3.2.6.2", $"{srcLabel}[{i}]: {val:F2} вне диапазона [-1000; 1000]."); arr = null; return false;
                }
                arr[i] = Math.Round(val, 2);
                if (arr[i] == 0.0) arr[i] = 0.0; // Страховка от бага "-0.00"
            }
            return true;
        }

        private void DGV_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (isUpdatingGrid || e.RowIndex != 0) return;
            string raw = e.FormattedValue?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(raw))
            {
                dgvEditable.Rows[0].ErrorText = "Ячейка не может быть пустой!"; return;
            }
            
            string valStr = raw.Replace(',', '.');
            if (!double.TryParse(valStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
            {
                dgvEditable.Rows[0].ErrorText = $"«{raw}» — не число.";
                Log($"Ячейка [{e.ColumnIndex}]: «{raw}» — не является числом.", true);
            }
            else if (val < -1000 || val > 1000)
            {
                dgvEditable.Rows[0].ErrorText = $"{val:F2} вне диапазона [-1000; 1000]";
                Log($"Ячейка [{e.ColumnIndex}]: {val:F2} вне ОДЗ [-1000; 1000].", true);
            }
            else
            {
                dgvEditable.Rows[0].ErrorText = "";
            }
        }
        private void PrepareManualInput()
        {
            using (Form dlg = new Form
            {
                Width = 400,
                Height = 180,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Ручной ввод массива",
                StartPosition = FormStartPosition.CenterParent,
                Font = new Font("Segoe UI", 9)
            })
            {
                var lblVals = new Label { Text = "Введите числа через пробел (от 1 до 20 элементов):", Left = 15, Top = 15, Width = 350 };
                var txtVals = new TextBox { Left = 15, Top = 35, Width = 350, Font = new Font("Segoe UI", 10), BorderStyle = BorderStyle.FixedSingle };

                var btnOk = new Button
                {
                    Text = "Применить",
                    Left = 130,
                    Top = 85,
                    Width = 110,
                    Height = 35,
                    BackColor = CLR_ACCENT,
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                var btnCancel = new Button
                {
                    Text = "Отмена",
                    Left = 250,
                    Top = 85,
                    Width = 110,
                    Height = 35,
                    FlatStyle = FlatStyle.Flat,
                    DialogResult = DialogResult.Cancel
                };

                btnOk.FlatAppearance.BorderSize = 0;

                btnOk.Click += (s, e) =>
                {
                    if (string.IsNullOrWhiteSpace(txtVals.Text))
                    {
                        ShowError("3.2.6.4", "Вы не ввели элементы массива. Поле пустое.");
                        ClearAllData();
                        dlg.DialogResult = DialogResult.Abort;
                        return;
                    }
                    dlg.DialogResult = DialogResult.OK;
                };

                dlg.Controls.AddRange(new Control[] { lblVals, txtVals, btnOk, btnCancel });
                dlg.AcceptButton = btnOk; dlg.CancelButton = btnCancel;

                if (dlg.ShowDialog() != DialogResult.OK) return;

                string[] parts = txtVals.Text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                int n = parts.Length;

                if (n < 1 || n > 20)
                {
                    ShowError("3.2.6.1", $"Некорректный размер массива: обнаружено {n} элементов. Должно быть от 1 до 20.");
                    ClearAllData();
                    return;
                }

                if (!TryParseElements(parts, n, "Элемент", out double[] arr))
                {
                    ClearAllData(); 
                    return;
                }

                InitializeNewData(arr, "Ручной ввод");
            }
        }

        private void GenerateRandom()
        {
            using (Form dlg = new Form
            {
                Width = 350,
                Height = 200,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Случайная генерация",
                StartPosition = FormStartPosition.CenterParent,
                Font = new Font("Segoe UI", 9)
            })
            {
                var chkRandomN = new CheckBox { Text = "Полностью случайный размер массива", Left = 15, Top = 20, Width = 300 };
                var lblN = new Label { Text = "Точный размер N [1..20]:", Left = 15, Top = 60, Width = 200 };
                var numN = new NumericUpDown { Left = 15, Top = 80, Width = 80, Minimum = 1, Maximum = 20, Value = 5 };

                chkRandomN.CheckedChanged += (s, e) => numN.Enabled = !chkRandomN.Checked;

                var btnOk = new Button
                {
                    Text = "Сгенерировать",
                    Left = 100,
                    Top = 120,
                    Width = 110,
                    Height = 35,
                    BackColor = CLR_ACCENT,
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                var btnCancel = new Button
                {
                    Text = "Отмена",
                    Left = 220,
                    Top = 120,
                    Width = 90,
                    Height = 35,
                    FlatStyle = FlatStyle.Flat,
                    DialogResult = DialogResult.Cancel
                };

                btnOk.FlatAppearance.BorderSize = 0;

                btnOk.Click += (s, e) =>
                {
                    if (!chkRandomN.Checked && string.IsNullOrWhiteSpace(numN.Text))
                    {
                        ShowError("3.2.6.1", "Поле размера N не может быть пустым.");
                        ClearAllData();
                        dlg.DialogResult = DialogResult.Abort;
                        return;
                    }
                    dlg.DialogResult = DialogResult.OK;
                };

                dlg.Controls.AddRange(new Control[] { chkRandomN, lblN, numN, btnOk, btnCancel });
                dlg.AcceptButton = btnOk; dlg.CancelButton = btnCancel;

                if (dlg.ShowDialog() != DialogResult.OK) return;

                int n = chkRandomN.Checked ? globalRandom.Next(1, 21) : (int)numN.Value;

                double[] arr = Enumerable.Range(0, n).Select(_ => Math.Round(globalRandom.NextDouble() * 2000 - 1000, 2)).ToArray();
                InitializeNewData(arr, "Случайная генерация");
            }
        }

        private void InitializeNewData(double[] arr, string method)
        {
            initialArray = arr.ToArray();
            editableArray = arr.ToArray();
            resultArray = null;

            inputMethod = method;
            isManuallyEdited = false;

            UpdateGrid(dgvInitial, initialArray);
            UpdateGrid(dgvEditable, editableArray);
            UpdateGrid(dgvResult, null);

            btnCalculate.Enabled = true;
            UpdateArrayInfo();
            Log($"Массив успешно загружен ({method}). Размер: {arr.Length}.");
        }
        private void UpdateGrid(DataGridView g, double[] arr)
        {
            isUpdatingGrid = true;
            g.Rows.Clear();
            g.Columns.Clear();
            if (arr == null || arr.Length == 0) { isUpdatingGrid = false; return; }

            g.ColumnCount = arr.Length;
            for (int i = 0; i < arr.Length; i++)
            {
                g.Columns[i].Width = 70;
                g.Columns[i].HeaderText = $"[{i}]";
                g.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            g.Rows.Add(arr.Select(x => (object)x.ToString("F2")).ToArray());
            isUpdatingGrid = false;
        }
        private struct Segment
        {
            public int Left, Right;
            public string LeftRole, RightRole;
            public double LeftVal, RightVal;
            public int[] ElementsToZero => Right - Left > 1 ? Enumerable.Range(Left + 1, Right - Left - 1).ToArray() : new int[0];
            public override string ToString() => ElementsToZero.Length > 0 ? $"{LeftRole}[{Left}]={LeftVal:F2} ↔ {RightRole}[{Right}]={RightVal:F2}" : $"{LeftRole}[{Left}] ↔ {RightRole}[{Right}] (смежные)";
        }

        private void ProcessArray()
{
    try
    {
        dgvEditable.EndEdit();
        if (editableArray == null || editableArray.Length != dgvEditable.ColumnCount)
        {
            editableArray = new double[dgvEditable.ColumnCount];
        }
        for (int i = 0; i < dgvEditable.ColumnCount; i++)
        {
            string raw = dgvEditable.Rows[0].Cells[i].Value?.ToString().Replace(',', '.') ?? "";
            
            if (!double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double val)) 
            { 
                ShowError("3.2.6.6", $"Ячейка [{i}]: нечисловое значение."); 
                return; 
            }
            if (val < -1000 || val > 1000) 
            { 
                ShowError("3.2.6.2", $"Ячейка [{i}]: {val:F2} вне [-1000; 1000]."); 
                return; 
            }
            editableArray[i] = Math.Round(val, 2);
            if (editableArray[i] == 0.0) editableArray[i] = 0.0; // Страховка от системного "-0.00"
        }

        resultArray = editableArray.ToArray();
        if (resultArray.Length < 3)
        {
            Log(resultArray.Length == 1 ? "В массиве 1 элемент — обнулять нечего." : "В массиве 2 элемента — обнулять нечего.");
            UpdateGrid(dgvResult, resultArray); 
            return;
        }

        double minVal = resultArray.Min();
        double maxVal = resultArray.Max();
        if (minVal == maxVal)
        {
            DialogResult dr = MessageBox.Show(
                $"Все элементы равны {minVal:F2}.\nОбнулить элементы [1..{resultArray.Length - 2}]?", 
                "Min = Max", 
                MessageBoxButtons.YesNoCancel, 
                MessageBoxIcon.Question
            );

            if (dr == DialogResult.Yes)
            {
                for (int i = 1; i < resultArray.Length - 1; i++) resultArray[i] = 0.0;
                Log($"Min=Max={minVal:F2}. Обнулены [1..{resultArray.Length - 2}].");
                UpdateGrid(dgvResult, resultArray);
            }
            else if (dr == DialogResult.No)
            {
                Log($"Min=Max={minVal:F2}. Отменено (оставлен оригинал).");
                UpdateGrid(dgvResult, resultArray);
            }
            else 
            {
                Log("Операция прервана пользователем.", true);
                UpdateGrid(dgvResult, null);
            }
            return;
        }
        var pos = resultArray.Select((v, i) => (v, i, role: v == minVal ? "Min" : v == maxVal ? "Max" : null)).Where(x => x.role != null).OrderBy(x => x.i).ToList();
        if (pos.Count(x => x.role == "Min") == 1 && pos.Count(x => x.role == "Max") == 1)
        {
            resultArray = ArrayProcessor.ProcessVariant9(editableArray);
            
            if (resultArray.SequenceEqual(editableArray))
            {
                Log("Min и Max стоят рядом. Нечего обнулять.");
            }
            else
            {
                Log($"Массив успешно обработан через класс ArrayProcessor. Элементы между {minVal:F2} и {maxVal:F2} обнулены.");
            }
        }
        else
        {
            var segsV1 = new List<Segment>();
            for (int k = 0; k < pos.Count - 1; k++)
                if (pos[k].role != pos[k + 1].role) segsV1.Add(new Segment { Left = pos[k].i, Right = pos[k + 1].i, LeftRole = pos[k].role, RightRole = pos[k + 1].role, LeftVal = pos[k].v, RightVal = pos[k + 1].v });

            var segsV3 = new List<Segment>();
            for (int i = 0; i < pos.Count; i++)
                for (int j = i + 1; j < pos.Count; j++)
                    if (pos[i].role != pos[j].role) segsV3.Add(new Segment { Left = pos[i].i, Right = pos[j].i, LeftRole = pos[i].role, RightRole = pos[j].role, LeftVal = pos[i].v, RightVal = pos[j].v });

            int fMin = pos.First(x => x.role == "Min").i; int lMin = pos.Last(x => x.role == "Min").i;
            int fMax = pos.First(x => x.role == "Max").i; int lMax = pos.Last(x => x.role == "Max").i;

            int d1 = lMax > fMin ? lMax - fMin : 0;
            int d2 = lMin > fMax ? lMin - fMax : 0;

            int wL = -1, wR = -1; string wLR = "", wRR = "";
            if (d1 >= d2 && d1 > 0) { wL = fMin; wR = lMax; wLR = "Min"; wRR = "Max"; }
            else if (d2 > d1 && d2 > 0) { wL = fMax; wR = lMin; wLR = "Max"; wRR = "Min"; }

            bool wOk = wL != -1 && (wR - wL > 1);

            int[] chosen = ShowDuplicateDialog(segsV1, segsV3, wOk, wL, wR, wLR, wRR, minVal, maxVal);
            
            if (chosen == null) 
            {
                Log("Операция прервана пользователем.", true);
                UpdateGrid(dgvResult, null);
                return; 
            }

            foreach (int idx in chosen) resultArray[idx] = 0.0;

            if (chosen.Length > 0)
            {
                Log($"Обнулено {chosen.Length} элементов из-за дубликатов Min/Max.");
            }
            else
            {
                Log("Дубликаты обработаны: нет подходящих элементов для обнуления между парами.");
            }
        }
        UpdateGrid(dgvResult, resultArray);
    }
    catch (Exception ex) 
    { 
        Log(ex.Message, true); 
    }
}

        private int[] ShowDuplicateDialog(List<Segment> v1, List<Segment> v3, bool wOk, int wL, int wR, string wLR, string wRR, double min, double max)
        {
            var validV3 = v3.Where(s => s.ElementsToZero.Length > 0).ToList();
            Form dlg = new Form { Width = 600, Text = "Неоднозначность: выбор сегмента", StartPosition = FormStartPosition.CenterParent, Font = new Font("Segoe UI", 9) };

            int y = 12;
            dlg.Controls.Add(new Label { Left = 12, Top = y, Width = 560, Text = $"Найдено несколько Min({min:F2}) или Max({max:F2}). Выберите область обнуления:" }); y += 30;

            var rb1 = new RadioButton { Left = 12, Top = y, Width = 560, Checked = true, Text = "Вариант 1: Все смежные сегменты (каждая соседняя пара Min↔Max)" }; dlg.Controls.Add(rb1); y += 25;
            var rb2 = new RadioButton { Left = 12, Top = y, Width = 560, Enabled = wOk, Text = wOk ? $"Вариант 2: Максимальный охват (от {wLR}[{wL}] до {wRR}[{wR}])" : "Вариант 2: Максимальный охват (Недоступно)" }; dlg.Controls.Add(rb2); y += 25;
            var rb3 = new RadioButton { Left = 12, Top = y, Width = 560, Text = "Вариант 3: Ручной выбор комбинаций" }; dlg.Controls.Add(rb3); y += 25;

            var clb = new CheckedListBox { Left = 30, Top = y, Width = 540, Height = 100, Enabled = false };
            foreach (var s in validV3) clb.Items.Add(s.ToString());
            dlg.Controls.Add(clb); y += 110;

            rb3.CheckedChanged += (s, e) => clb.Enabled = rb3.Checked;

            var btnOk = new Button { Text = "ОК", Left = 340, Top = y, Width = 110, DialogResult = DialogResult.OK, BackColor = CLR_ACCENT, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            var btnCancel = new Button { Text = "Отмена", Left = 460, Top = y, Width = 110, DialogResult = DialogResult.Cancel, FlatStyle = FlatStyle.Flat };

            dlg.Controls.AddRange(new Control[] { btnOk, btnCancel });
            dlg.AcceptButton = btnOk; dlg.CancelButton = btnCancel;
            dlg.ClientSize = new Size(600, y + 40);

            if (dlg.ShowDialog() != DialogResult.OK) return null;

            var res = new HashSet<int>();
            if (rb1.Checked)
            {
                foreach (var s in v1) foreach (int i in s.ElementsToZero) res.Add(i);
            }
            else if (rb2.Checked)
            {
                for (int i = wL + 1; i < wR; i++) res.Add(i);
            }
            else if (rb3.Checked)
            {
                foreach (int idx in clb.CheckedIndices) foreach (int i in validV3[idx].ElementsToZero) res.Add(i);
            }
            return res.ToArray();
        }
        private void LoadTxt()
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "Text Files|*.txt" };
            if (ofd.ShowDialog() != DialogResult.OK) return;
            try
            {
                if (!File.Exists(ofd.FileName)) { ShowError("3.2.6.5", "Файл не найден."); ClearAllData(); return; }
                string[] lines = File.ReadAllLines(ofd.FileName, Encoding.UTF8).Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
                if (lines.Length < 2) { ShowError("3.2.6.3", "Файл должен содержать минимум 2 непустые строки (N и элементы)."); ClearAllData(); return; }
                if (!int.TryParse(lines[0].Trim(), out int n) || n < 1 || n > 20)
                {
                    ShowError("3.2.6.1", "Строка 1: N вне диапазона [1;20]."); ClearAllData(); return;
                }
                string[] parts = lines[1].Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                if (!TryParseElements(parts, n, "Элемент", out double[] arr)) { ClearAllData(); return; }

                InitializeNewData(arr, "Загрузка из .txt");
            }
            catch (Exception ex) { ClearAllData(); ShowError("3.2.6.3", "Ошибка чтения: " + ex.Message); }
        }

        private void SaveTxt()
        {
            if (resultArray == null)
            {
                MessageBox.Show("Сначала выполните обнуление! Сохранять можно только итоговый результат.", "Ошибка сохранения", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            SaveFileDialog sfd = new SaveFileDialog { Filter = "Text Files|*.txt" };
            if (sfd.ShowDialog() != DialogResult.OK) return;
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine(resultArray.Length.ToString());
                sb.AppendLine(string.Join(" ", resultArray.Select(x => x.ToString("F2"))));
                sb.AppendLine();
                sb.AppendLine($"Способ ввода оригинала: {inputMethod}");
                sb.AppendLine($"Отредактирован вручную перед обработкой: {(isManuallyEdited ? "Да" : "Нет")}");
                File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                Log("Итоговый результат сохранён в .txt.");
            }
            catch (Exception ex) { ShowError("3.2.6.7", ex.Message); }
        }

        private void LoadExcel()
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel Files|*.xlsx" };
            if (ofd.ShowDialog() != DialogResult.OK) return;
            try
            {
                using (var wb = new XLWorkbook(ofd.FileName))
                {
                    var ws = wb.Worksheet(1);
                    string nStr = ws.Cell(1, 1).GetValue<string>();

                    if (string.IsNullOrWhiteSpace(nStr) || !int.TryParse(nStr, out int n) || n < 1 || n > 20)
                    {
                        ShowError("3.2.6.1", "Ячейка A1: пусто или N вне диапазона [1;20]."); ClearAllData(); return;
                    }

                    string[] parts = new string[n];
                    for (int i = 0; i < n; i++)
                    {
                        string val = ws.Cell(i + 2, 1).GetValue<string>();
                        if (string.IsNullOrWhiteSpace(val))
                        {
                            ShowError("3.2.6.9", $"Обнаружена пустая ячейка A{i + 2}. Ввод прерван."); ClearAllData(); return;
                        }
                        parts[i] = val;
                    }

                    if (!TryParseElements(parts, n, "Ячейка A", out double[] arr)) { ClearAllData(); return; }

                    InitializeNewData(arr, "Импорт из Excel");
                }
            }
            catch (Exception ex) { ClearAllData(); ShowError("3.2.6.8", ex.Message); }
        }

        private void SaveExcel()
        {
            if (resultArray == null)
            {
                MessageBox.Show("Сначала выполните обнуление! Сохранять можно только итоговый результат.", "Ошибка сохранения", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Files|*.xlsx" };
            if (sfd.ShowDialog() != DialogResult.OK) return;
            try
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Результат");
                    ws.Cell(1, 1).Value = resultArray.Length;
                    for (int i = 0; i < resultArray.Length; i++) ws.Cell(i + 2, 1).Value = resultArray[i];
                    ws.Cell(1, 2).Value = $"Способ ввода оригинала: {inputMethod}";
                    ws.Cell(2, 2).Value = $"Отредактирован вручную перед обработкой: {(isManuallyEdited ? "Да" : "Нет")}";
                    ws.Cell(1, 2).Style.Font.Bold = true;
                    ws.Cell(2, 2).Style.Font.Bold = true;
                    ws.Columns().AdjustToContents();
                    wb.SaveAs(sfd.FileName);
                }
                Log("Итоговый результат экспортирован в Excel.");
            }
            catch (Exception ex) { ShowError("3.2.6.7", ex.Message); }
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}