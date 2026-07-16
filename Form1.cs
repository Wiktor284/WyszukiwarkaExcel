using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelDataReader;

namespace WyszukiwarkaExcel
{
    public partial class Form1 : Form
    {
        private DataTable _tabelaOryginalna;
        private DataTable _tabelaFiltrowana;

        // Elementy interfejsu
        private Button btnWczytaj;
        private Label lblStatus;
        private TextBox txtSzukaj;
        private DataGridView dgvWyniki;

        public Form1()
        {
            StworzInterfejs();
            
            //wymagane dla biblioteki ExcelDataReader w środowisku .NET
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        //metoda tworzaca wyglad okna programu
        private void StworzInterfejs()
        {
            this.Text = "Lokalna Wyszukiwarka Osób - C# Desktop App";
            this.Width = 950;
            this.Height = 650;
            this.StartPosition = FormStartPosition.CenterScreen;

            //panel gorny (szary kontener na przyciski i pole szukania)
            Panel panelTop = new Panel { Dock = DockStyle.Top, Height = 110, BackColor = System.Drawing.Color.FromArgb(240, 243, 244) };

            btnWczytaj = new Button 
            { 
                Text = "Załaduj plik Excel", 
                Location = new System.Drawing.Point(15, 15), 
                Size = new System.Drawing.Size(180, 35),
                BackColor = System.Drawing.Color.FromArgb(41, 128, 185),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new System.Drawing.Font("Segoe UI", 9.5f, System.Drawing.FontStyle.Bold)
            };
            btnWczytaj.FlatAppearance.BorderSize = 0;
            btnWczytaj.Click += BtnWczytaj_Click;

            lblStatus = new Label 
            { 
                Text = "Brak danych. Kliknij przycisk obok i wybierz plik .xlsx lub .xls", 
                Location = new System.Drawing.Point(210, 22), 
                Size = new System.Drawing.Size(700, 25), 
                Font = new System.Drawing.Font("Segoe UI", 9.5f),
                ForeColor = System.Drawing.Color.DimGray 
            };

            txtSzukaj = new TextBox 
            { 
                PlaceholderText = "🔍 Wpisz dowolne dane, aby wyszukać...", 
                Location = new System.Drawing.Point(15, 65), 
                Size = new System.Drawing.Size(900, 30), 
                Font = new System.Drawing.Font("Segoe UI", 11),
                Enabled = false 
            };
            txtSzukaj.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            txtSzukaj.TextChanged += TxtSzukaj_TextChanged;

            panelTop.Controls.Add(btnWczytaj);
            panelTop.Controls.Add(lblStatus);
            panelTop.Controls.Add(txtSzukaj);

            //tabela z wynikami
            dgvWyniki = new DataGridView 
            { 
                Dock = DockStyle.Fill, 
                ReadOnly = true, 
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.None,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = System.Drawing.Color.FromArgb(245, 247, 250) }
            };

            this.Controls.Add(dgvWyniki);
            this.Controls.Add(panelTop);
            
            //wymuszzenie poprawnego ulozenia kontrolek
            panelTop.SendToBack(); 
            dgvWyniki.BringToFront();
        }

        //logika recznego wczytywania pliku excel
        private void BtnWczytaj_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Pliki Excel|*.xlsx;*.xls";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    WczytajExcel(ofd.FileName);
                }
            }
        }

        //metoda do ladowania wskazanego pliku
        private void WczytajExcel(string sciezkaDoPliku)
        {
            try
            {
                //bezpieczne otwarcie pliku (nawet jesli ktos ma go akurat otwartego w excelu)
                using (var stream = File.Open(sciezkaDoPliku, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });

                        if (result.Tables.Count > 0)
                        {
                            _tabelaOryginalna = result.Tables[0];
                            dgvWyniki.DataSource = _tabelaOryginalna;
                            lblStatus.Text = $"Załadowano: {Path.GetFileName(sciezkaDoPliku)} (Wierszy: {_tabelaOryginalna.Rows.Count})";
                            lblStatus.ForeColor = System.Drawing.Color.DarkGreen;
                            txtSzukaj.Enabled = true;
                            txtSzukaj.Text = "";
                            txtSzukaj.Focus();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Wystąpił błąd podczas odczytu pliku:\n{ex.Message}", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //logika wyszukiwania w czasie rzeczywistym
        private void TxtSzukaj_TextChanged(object sender, EventArgs e)
        {
            if (_tabelaOryginalna == null) return;

            // Usuwamy polskie znaki z wpisanej frazy i przycinamy spacje
            string fraza = UsunPolskieZnaki(txtSzukaj.Text.Trim());
            
            if (string.IsNullOrEmpty(fraza))
            {
                dgvWyniki.DataSource = _tabelaOryginalna;
                return;
            }

            _tabelaFiltrowana = _tabelaOryginalna.Clone();

            foreach (DataRow wiersz in _tabelaOryginalna.Rows)
            {
                bool znaleziono = false;
                foreach (var komorka in wiersz.ItemArray)
                {
                    if (komorka != null)
                    {
                        // Usuwamy polskie znaki z tekstu w komórce przed porównaniem
                        string tekstKomorki = UsunPolskieZnaki(komorka.ToString());

                        if (tekstKomorki.Contains(fraza))
                        {
                            znaleziono = true;
                            break; 
                        }
                    }
                }

                if (znaleziono)
                {
                    _tabelaFiltrowana.ImportRow(wiersz);
                }
            }

            dgvWyniki.DataSource = _tabelaFiltrowana;
        }

        private string UsunPolskieZnaki(string tekst)
        {
            if (string.IsNullOrEmpty(tekst)) return "";

            var mapa = new (char polski, char zwykly)[]
            {
                ('ą', 'a'), ('ć', 'c'), ('ę', 'e'), ('ł', 'l'), ('ń', 'n'), 
                ('ó', 'o'), ('ś', 's'), ('ź', 'z'), ('ż', 'z')
            };

            string wynik = tekst.ToLower();

            foreach (var para in mapa)
            {
                wynik = wynik.Replace(para.polski, para.zwykly);
            }

            return wynik;
        }
    }
}