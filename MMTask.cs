using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Globalization; 

namespace Inventoris
{
	public class MMTask : System.Windows.Forms.Form
	{
		private string strKoneksi = "Data Source=.\\SQLEXPRESS; Initial Catalog=MM_DB; User ID=sa; Password=password123;";
		
		private System.Windows.Forms.Panel panelSidebar;
		private System.Windows.Forms.Panel panelHeader;
		private System.Windows.Forms.Panel panelContent;
		private System.Windows.Forms.DataGrid dgTransaksi;
		
		private System.Windows.Forms.TextBox txtCari;
		private System.Windows.Forms.TextBox txtNominal; 
		private System.Windows.Forms.ComboBox cmbOperator;
		private System.Windows.Forms.TextBox txtFilterKode; 
		private System.Windows.Forms.TextBox txtFilterKodeRekAju; 
		private System.Windows.Forms.Button btnCari;
		private System.Windows.Forms.Button btnRefresh; 
		private System.Windows.Forms.Button btnExport;
		private System.Windows.Forms.Label labelCari;
		private System.Windows.Forms.Label lblNominal;
		private System.Windows.Forms.Label labelKode; 
		private System.Windows.Forms.DateTimePicker dtpMulai;
		private System.Windows.Forms.DateTimePicker dtpSelesai;
		private System.Windows.Forms.CheckBox cbPakaiTgl;
		private System.Windows.Forms.ComboBox cmbJenis; 
		private System.Windows.Forms.Label lblJenis;

		private System.Windows.Forms.Label labelJudul;
		private System.Windows.Forms.Label labelInfoFilter; 
		private System.Windows.Forms.Label labelN;
		private System.Windows.Forms.Label labelK;
		private System.Windows.Forms.Label labelKodeRekAju;
		private System.Windows.Forms.Label lblStatus;
		private System.Windows.Forms.Label lblTotalNilai; 
		private System.Windows.Forms.TextBox txtDetailKet;
		private System.Windows.Forms.TextBox txtDetailNilai;

		public MMTask() { InitializeComponent(); }

		private void txtNominal_Leave(object sender, EventArgs e)
		{
			try 
			{
				if (txtNominal.Text.Trim() != "") 
				{
					double dbl = double.Parse(txtNominal.Text.Replace(".", "").Replace(",", ""));
					txtNominal.Text = dbl.ToString("N0", new CultureInfo("id-ID"));
				}
			} 
			catch { }
		}

		private string SafeXml(object input) 
		{
			if (input == DBNull.Value || input == null) return "";
			string value = input.ToString();
			return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
		}

		private string GetNamaKurs(string kodeKurs)
		{
			try 
			{
				SqlConnection conn = new SqlConnection(strKoneksi);
				string sql = "SELECT NAMA_KURS FROM TB_KURS WHERE KODE_KURS = '" + kodeKurs + "'";
				SqlCommand cmd = new SqlCommand(sql, conn);
				conn.Open();
				object result = cmd.ExecuteScalar();
				conn.Close();
				
				if (result != null && result != DBNull.Value)
				{
					return result.ToString() + " (" + kodeKurs + ")";
				}
				return kodeKurs;
			} 
			catch 
			{
				return kodeKurs;
			}
		}

		private void AturKolomDataGrid()
		{
			DataGridTableStyle ts = new DataGridTableStyle();
			ts.MappingName = "tb_gabungan";
			ts.AlternatingBackColor = Color.FromArgb(240, 248, 255); // Light blue alternating
			ts.BackColor = Color.White;
			ts.ForeColor = Color.FromArgb(33, 37, 41); // Dark text
			ts.GridLineColor = Color.FromArgb(229, 229, 229); // Light gray grid
			ts.HeaderBackColor = Color.FromArgb(25, 118, 210); // Blue header
			ts.HeaderForeColor = Color.White;
			ts.HeaderFont = new Font("Segoe UI", 9, FontStyle.Bold);
			ts.PreferredRowHeight = 25;
			ts.RowHeaderWidth = 0;
			
			// Periksa kolom yang tersedia
			DataTable dt = dgTransaksi.DataSource as DataTable;
			if (dt != null)
			{
				// Menyiapkan array dinamis
				ArrayList mNames = new ArrayList();
				ArrayList hTexts = new ArrayList();
				ArrayList widths = new ArrayList();
				
				// Kolom dasar yang selalu ditampilkan
				string[] basicNames = {"nomor", "tanggal", "kode_count", "nilai", "KeteranganHeader", "KeteranganDetail", "jenis", "status_text", "no_aju", "kode_kurs"};
				string[] basicTexts = {"NO. TRANSAKSI", "TANGGAL", "KODE AKUN", "NILAI (IDR)", "KET. HEADER", "KET. DETAIL", "JENIS", "STATUS", "NO. AJU", "KURS"};
				int[] basicWidths = {120, 75, 70, 100, 160, 160, 60, 90, 90, 50};
				
				// Kolom tambahan yang tersedia
				string[] extraNames = {"NAMA_KURS", "NILAI_KURS", "nilai_konversi", "bank", "ket_transfer", "kode_rekaju", "ket_Status", "pengaju", "id_proses", "tgl_proses"};
				string[] extraTexts = {"NAMA KURS", "NILAI KURS", "NILAI KONVERSI", "BANK", "KET. TRANSFER", "KODE REK AJU", "KET. STATUS", "PENGAJU", "ID PROSES", "TGL PROSES"};
				int[] extraWidths = {80, 80, 100, 80, 120, 80, 100, 80, 80, 80};
				
				// Memproses kolom utama
				for(int i=0; i<basicNames.Length; i++) 
				{
					if(dt.Columns.Contains(basicNames[i])) 
					{
						mNames.Add(basicNames[i]);
						hTexts.Add(basicTexts[i]);
						widths.Add(basicWidths[i]);
					}
				}
				
				// Memproses kolom tambahan
				for(int i=0; i<extraNames.Length; i++) 
				{
					try 
					{
						if(dt.Columns.Contains(extraNames[i])) 
						{
							mNames.Add(extraNames[i]);
							hTexts.Add(extraTexts[i]);
							widths.Add(extraWidths[i]);
						}
					}
					catch 
					{
						// Skip kolom yang bermasalah
						continue;
					}
				}
				
				// Mengatur kolom DataGrid
				for(int i=0; i<mNames.Count; i++) 
				{
					DataGridTextBoxColumn col = new DataGridTextBoxColumn();
					col.MappingName = mNames[i].ToString();
					col.HeaderText = hTexts[i].ToString();
					col.Width = (int)widths[i];
					col.NullText = "";
					
					// Set alignment based on column type
					if(mNames[i].ToString() == "nilai" || mNames[i].ToString() == "nilai_konversi" || mNames[i].ToString() == "NILAI_KURS") { 
						col.Alignment = HorizontalAlignment.Right; 
						col.Format = "N0"; 
					}
					else if(mNames[i].ToString() == "nomor" || mNames[i].ToString() == "kode_kurs" || mNames[i].ToString() == "kode_count" || mNames[i].ToString() == "kode_rekaju" || mNames[i].ToString() == "no_aju" || mNames[i].ToString() == "pengaju" || mNames[i].ToString() == "id_proses") {
						col.Alignment = HorizontalAlignment.Center;
					}
					else {
						col.Alignment = HorizontalAlignment.Left;
					}
					
					// Set font for better readability
					col.TextBox.Font = new Font("Segoe UI", 9);
					
					ts.GridColumnStyles.Add(col);
				}
			}
			
			dgTransaksi.TableStyles.Clear();
			dgTransaksi.TableStyles.Add(ts);
			
			// Menambahkan styling DataGrid
			dgTransaksi.BackgroundColor = Color.FromArgb(248, 251, 253); // Very light blue
			dgTransaksi.BorderStyle = BorderStyle.FixedSingle;
			dgTransaksi.GridLineColor = Color.FromArgb(229, 229, 229);
			dgTransaksi.HeaderFont = new Font("Segoe UI", 9, FontStyle.Bold);
			dgTransaksi.Font = new Font("Segoe UI", 9);
			dgTransaksi.ForeColor = Color.FromArgb(33, 37, 41);
			dgTransaksi.PreferredColumnWidth = 100;
			dgTransaksi.PreferredRowHeight = 25;
		}

		private void SinkronisasiData() 
		{
			SqlConnection conn = new SqlConnection(strKoneksi);
			string infoText = "Filter Aktif: "; 
			bool filtered = false;
			try 
			{
				string baseQuery = @"
    -- Data dari tabel transaksi (header)
    SELECT 
        h.nomor, 
        h.bank,             -- 1. BANK
        h.tanggal,          -- 2. TANGGAL
        h.jenis,            -- 3. JENIS
        CASE h.status 
            WHEN 0 THEN 'BELUM PROSES' 
            WHEN 1 THEN 'DITERIMA' 
            WHEN 2 THEN 'DITOLAK' 
            ELSE '' 
        END AS status_text, -- 4. STATUS
        '' AS kode_count,   -- 5. KODE AKUN (kosong untuk header)
        k.NAMA_KURS,        -- 6. KURS
        0 AS nilai,         -- 7. NILAI (kosong untuk header)
        k.NILAI_KURS,       -- 8. NILAI KURS
        h.nilai_konversi,   -- 9. KONVERSI
        h.kode_rek,         -- 10. BANK (Kode Rekening)
        h.ket_Status,       -- 11. KET. STATUS
        h.no_aju,           -- 12. NO. AJU
        h.id AS pengaju,    -- 13. PENGAJU
        h.id_proses,        -- 14. DIPROSES OLEH
        ISNULL(CAST(h.tgl_proses AS VARCHAR), '') AS tgl_proses_str, -- 15. TGL PROSES sebagai string
        h.keterangan AS KeteranganHeader, -- 16
        '' AS KeteranganDetail, -- 17 (kosong untuk header)
        h.ket_transfer,     -- 18
        h.kode_rekaju,      -- 19
        h.kode_kurs         -- 20. KODE KURS
    FROM tb_transaksikasbankaju h
    LEFT JOIN TB_KURS k ON h.kode_kurs = k.KODE_KURS
    
    UNION ALL
    
    -- Data dari tabel detail transaksi
    SELECT 
        h.nomor, 
        h.bank,             -- 1. BANK
        h.tanggal,          -- 2. TANGGAL
        h.jenis,            -- 3. JENIS
        CASE h.status 
            WHEN 0 THEN 'BELUM PROSES' 
            WHEN 1 THEN 'DITERIMA' 
            WHEN 2 THEN 'DITOLAK' 
            ELSE '' 
        END AS status_text, -- 4. STATUS
        d.kode_count,       -- 5. KODE AKUN
        k.NAMA_KURS,        -- 6. KURS
        d.nilai,            -- 7. NILAI
        k.NILAI_KURS,       -- 8. NILAI KURS
        h.nilai_konversi,   -- 9. KONVERSI
        h.kode_rek,         -- 10. BANK (Kode Rekening)
        h.ket_Status,       -- 11. KET. STATUS
        h.no_aju,           -- 12. NO. AJU
        h.id AS pengaju,    -- 13. PENGAJU
        h.id_proses,        -- 14. DIPROSES OLEH
        ISNULL(CAST(h.tgl_proses AS VARCHAR), '') AS tgl_proses_str, -- 15. TGL PROSES sebagai string
        h.keterangan AS KeteranganHeader, -- 16
        d.keterangan AS KeteranganDetail, -- 17
        h.ket_transfer,     -- 18
        h.kode_rekaju,      -- 19
        h.kode_kurs         -- 20. KODE KURS
    FROM tb_transaksikasbankaju h
    INNER JOIN tb_detailtransaksikasbankaju d ON h.nomor = d.nomor
    LEFT JOIN TB_KURS k ON h.kode_kurs = k.KODE_KURS";
				
				string sql = "SELECT * FROM (" + baseQuery + ") AS combined_data WHERE 1=1";
				
				if (txtCari.Text.Trim() != "") 
				{
					sql += " AND (nomor LIKE '%" + txtCari.Text + "%' OR KeteranganHeader LIKE '%" + txtCari.Text + "%' OR KeteranganDetail LIKE '%" + txtCari.Text + "%' OR kode_count LIKE '%" + txtCari.Text + "%' OR no_aju LIKE '%" + txtCari.Text + "%' OR bank LIKE '%" + txtCari.Text + "%' OR kode_kurs LIKE '%" + txtCari.Text + "%' OR kode_rekaju LIKE '%" + txtCari.Text + "%' OR ket_Status LIKE '%" + txtCari.Text + "%' OR status_text LIKE '%" + txtCari.Text + "%')";
					infoText += "[Cari: " + txtCari.Text + "] "; filtered = true;
				}
				if (txtNominal.Text.Trim() != "") 
				{
					string cleanNominal = txtNominal.Text.Replace(".", "");
					sql += " AND nilai " + cmbOperator.SelectedItem.ToString() + " " + cleanNominal;
					infoText += "[Nilai " + cmbOperator.SelectedItem.ToString() + " " + txtNominal.Text + "] "; filtered = true;
				}
				if (txtFilterKode.Text.Trim() != "") 
				{
					sql += " AND kode_count LIKE '%" + txtFilterKode.Text + "%'";
					infoText += "[Kode: " + txtFilterKode.Text + "] "; filtered = true;
				}
				if (txtFilterKodeRekAju.Text.Trim() != "") 
				{
					sql += " AND kode_rekaju LIKE '%" + txtFilterKodeRekAju.Text + "%'";
					infoText += "[Kode Rek Aju: " + txtFilterKodeRekAju.Text + "] "; filtered = true;
				}
				if (cmbJenis.SelectedIndex > 0) 
				{
					sql += " AND jenis = '" + cmbJenis.SelectedItem.ToString() + "'";
					infoText += "[Jenis: " + cmbJenis.SelectedItem.ToString() + "] "; filtered = true;
				}
				if (cbPakaiTgl.Checked) 
				{
					sql += " AND (tanggal >= '" + dtpMulai.Value.ToString("yyyy-MM-dd") + "' AND tanggal <= '" + dtpSelesai.Value.ToString("yyyy-MM-dd") + "')";
					infoText += "[Tgl: " + dtpMulai.Value.ToString("dd/MM/yy") + "-" + dtpSelesai.Value.ToString("dd/MM/yy") + "] "; filtered = true;
				}

				if (!filtered) infoText = "Filter Aktif: SEMUA DATA";

				SqlDataAdapter da = new SqlDataAdapter(sql, conn);
				DataTable dt = new DataTable("tb_gabungan");
				da.Fill(dt);

					// Menambahkan kolom yang diperlukan
				string[] requiredColumns = {"tgl_format", "Ket_Header", "Ket_Detail", "kode_rekaju", "ket_Status", "tgl_proses"};
				foreach(string col in requiredColumns) {
					if (!dt.Columns.Contains(col)) {
						dt.Columns.Add(col, typeof(string));
					}
				}
				
				long total = 0;
				foreach(DataRow row in dt.Rows) 
				{
					if(row["tanggal"] != DBNull.Value) row["tgl_format"] = Convert.ToDateTime(row["tanggal"]).ToString("dd MMM yy");
					if(row["nilai"] != DBNull.Value) total += Convert.ToInt64(row["nilai"]);
					
					// Isi kolom yang kosong
					if(row["Ket_Header"] == DBNull.Value) row["Ket_Header"] = "";
					if(row["Ket_Detail"] == DBNull.Value) row["Ket_Detail"] = "";
					if(row["kode_rekaju"] == DBNull.Value) row["kode_rekaju"] = "";
					if(row["ket_Status"] == DBNull.Value) row["ket_Status"] = "";
					
					// Format tanggal proses
					// Menggunakan kolom string dari SQL
					string rawTglProses = row["tgl_proses_str"].ToString();

					if (rawTglProses == null || rawTglProses == "" || row["status_text"].ToString() == "BELUM PROSES") 
					{
						row["tgl_proses"] = "-"; // Menampilkan strip jika belum ada tanggal
					} 
					else 
					{
						try 
						{
							DateTime tgl = Convert.ToDateTime(rawTglProses);
							row["tgl_proses"] = tgl.ToString("dd-MM-yyyy");
						} 
						catch 
						{
							row["tgl_proses"] = "-"; 
						}
					}
				}

				dgTransaksi.DataSource = dt;
				AturKolomDataGrid();
				lblTotalNilai.Text = "TOTAL NILAI: Rp " + total.ToString("N0", new CultureInfo("id-ID"));
				labelInfoFilter.Text = infoText;
				
				// Info kolom untuk debugging
				ArrayList columnNames = new ArrayList();
				foreach(DataColumn col in dt.Columns)
				{
					columnNames.Add(col.ColumnName);
				}
				string kolomInfo = string.Join(", ", (string[])columnNames.ToArray(typeof(string)));
				
				// Memeriksa ketersediaan kolom penting
				bool hasKodeRekaju = dt.Columns.Contains("kode_rekaju");
				bool hasKetStatus = dt.Columns.Contains("ket_Status");
				bool hasStatusLengkap = dt.Columns.Contains("status_lengkap");
				
				lblStatus.Text = "Total Data: " + dt.Rows.Count + " baris";
			} 
			catch (Exception ex) { MessageBox.Show("Gagal Sinkron: " + ex.Message); }
			finally { conn.Close(); }
		}

		private void btnExport_Click(object sender, EventArgs e)
		{
			SaveFileDialog save = new SaveFileDialog();
			save.Filter = "Excel XML (*.xls)|*.xls"; 
			save.FileName = "Laporan_Export_" + DateTime.Now.ToString("ddMMyy_HHmm") + ".xls";

			if (save.ShowDialog() == DialogResult.OK) 
			{
				try 
				{
					DataTable dt = (DataTable)dgTransaksi.DataSource;
					if (dt == null) return;
					// Mengkonfigurasi kolom Excel
					string[] allColumnNames = {"nomor", "tanggal", "kode_count", "nilai", "KeteranganHeader", "KeteranganDetail", "jenis", "status_text", "no_aju", "kode_kurs", "NAMA_KURS", "NILAI_KURS", "nilai_konversi", "bank", "ket_transfer", "kode_rekaju", "ket_Status", "pengaju", "id_proses", "tgl_proses"};
					string[] allHeaderTexts = {"NO. TRANSAKSI", "TANGGAL", "KODE AKUN", "NILAI (IDR)", "KET. HEADER", "KET. DETAIL", "JENIS", "STATUS", "NO. AJU", "KURS", "NAMA KURS", "NILAI KURS", "NILAI KONVERSI", "BANK", "KET. TRANSFER", "KODE REK AJU", "KET. STATUS", "PENGAJU", "ID PROSES", "TGL PROSES"};

					StringWriter sw = new StringWriter();
					
					sw.WriteLine("<?xml version=\"1.0\"?>");
					sw.WriteLine("<?mso-application progid=\"Excel.Sheet\"?>");
					sw.WriteLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
					sw.WriteLine(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
					sw.WriteLine(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"");
					sw.WriteLine(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");

					sw.WriteLine(" <Styles>");
					sw.WriteLine("  <Style ss:ID=\"Default\" ss:Name=\"Normal\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sTitle\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"22\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#1565C0\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Center\"/></Style>");
					sw.WriteLine("  <Style ss:ID=\"sSubtitle\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"10\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#546E7A\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Center\"/></Style>");
					sw.WriteLine("  <Style ss:ID=\"sHeaderBlue\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#1976D2\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"2\" ss:Color=\"#0D47A1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sHeaderGreen\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#388E3C\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"2\" ss:Color=\"#1B5E20\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sHeaderOrange\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#F57C00\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"2\" ss:Color=\"#E65100\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sHeaderPurple\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#7B1FA2\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"2\" ss:Color=\"#4A148C\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sHeaderRed\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#D32F2F\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"2\" ss:Color=\"#B71C1C\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sDataBlue\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Color=\"#1565C0\"/><Alignment ss:Vertical=\"Center\"/><Interior ss:Color=\"#E3F2FD\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sDataGreen\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Color=\"#2E7D32\"/><Alignment ss:Vertical=\"Center\"/><Interior ss:Color=\"#E8F5E8\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sDataOrange\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Color=\"#EF6C00\"/><Alignment ss:Vertical=\"Center\"/><Interior ss:Color=\"#FFF3E0\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sDataPurple\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Color=\"#6A1B9A\"/><Alignment ss:Vertical=\"Center\"/><Interior ss:Color=\"#F3E5F5\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sDataRed\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Color=\"#C62828\"/><Alignment ss:Vertical=\"Center\"/><Interior ss:Color=\"#FFEBEE\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sNumber\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"11\" ss:Bold=\"1\" ss:Color=\"#D84315\"/><Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Center\"/><NumberFormat ss:Format=\"#,##0\"/><Interior ss:Color=\"#FFF8E1\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sTotal\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"14\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#00695C\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Right\" ss:Vertical=\"Center\"/><NumberFormat ss:Format=\"#,##0\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"3\" ss:Color=\"#004D40\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sTotalCentered\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"14\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#00695C\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><NumberFormat ss:Format=\"#,##0\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"3\" ss:Color=\"#004D40\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sTotalLabel\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"14\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#00695C\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"3\" ss:Color=\"#004D40\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#E0E0E0\"/></Borders></Style>");
					sw.WriteLine("  <Style ss:ID=\"sFilterInfo\"><Font ss:FontName=\"Segoe UI\" ss:Size=\"10\" ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#FF6F00\" ss:Pattern=\"Solid\"/><Alignment ss:Horizontal=\"Center\"/></Style>");
					sw.WriteLine(" </Styles>");

					sw.WriteLine(" <Worksheet ss:Name=\"Laporan\">");
					sw.WriteLine("  <Table>");
					
					// Column widths - gunakan lebar default untuk 21 kolom
					ArrayList widths = new ArrayList();
					int[] defaultWidths = {120, 75, 70, 100, 160, 160, 60, 90, 90, 50, 80, 80, 60, 100, 80, 120, 80, 100, 80, 80, 80};
					for(int i=0; i<allColumnNames.Length; i++) {
						if(dt.Columns.Contains(allColumnNames[i])) {
							widths.Add(defaultWidths[i]);
						}
					}

					for(int i=0; i<widths.Count; i++) 
					{
						sw.WriteLine("   <Column ss:Width=\"" + widths[i] + "\"/>");
					}

					// Menampilkan judul laporan
					sw.WriteLine("   <Row ss:Height=\"30\"><Cell ss:MergeAcross=\"" + (widths.Count-1) + "\" ss:StyleID=\"sTitle\"><Data ss:Type=\"String\">🏢 MM GALLERI</Data></Cell></Row>");
					sw.WriteLine("   <Row ss:Height=\"30\"><Cell ss:MergeAcross=\"" + (widths.Count-1) + "\" ss:StyleID=\"sTitle\"><Data ss:Type=\"String\">📊 LAPORAN PENGAJUAN TRANSAKSI KAS BANK</Data></Cell></Row>");
					sw.WriteLine("   <Row ss:Height=\"10\"></Row>");

					// Menampilkan info filter dan waktu
					string filterInfo = labelInfoFilter.Text.Replace("Filter Aktif: ", "");
					string detailFilter = "";
            
					if (filterInfo == "SEMUA DATA") 
					{
						detailFilter = "Menampilkan semua data transaksi";
					} 
					else 
					{
						if (txtCari.Text.Trim() != "") 
						{
							detailFilter += "Pencarian untuk: '" + txtCari.Text + "' ";
						}
						if (txtNominal.Text.Trim() != "") 
						{
							string nominalText = txtNominal.Text.Replace(".", "");
							string operatorText = cmbOperator.SelectedItem.ToString();
							string operatorIndo = "";
                    
							switch(operatorText) 
							{
								case ">": operatorIndo = "lebih dari"; break;
								case ">=": operatorIndo = "lebih dari atau sama dengan"; break;
								case "=": operatorIndo = "sama dengan"; break;
								case "<": operatorIndo = "kurang dari"; break;
								case "<=": operatorIndo = "kurang dari atau sama dengan"; break;
							}
                    
							detailFilter += "Nilai " + operatorIndo + " Rp " + nominalText + " ";
						}
						if (txtFilterKode.Text.Trim() != "") 
						{
							detailFilter += "Kode akun yang mengandung: '" + txtFilterKode.Text + "' ";
						}
						if (txtFilterKodeRekAju.Text.Trim() != "") 
						{
							detailFilter += "Kode rek aju yang mengandung: '" + txtFilterKodeRekAju.Text + "' ";
						}
						if (cmbJenis.SelectedIndex > 0) 
						{
							detailFilter += "Jenis transaksi: " + cmbJenis.SelectedItem + " ";
						}
						if (cbPakaiTgl.Checked) 
						{
							detailFilter += "Periode transaksi: " + dtpMulai.Value.ToString("dd MMM yyyy") + " hingga " + dtpSelesai.Value.ToString("dd MMM yyyy") + " ";
						}
						// Menampilkan info filter
						sw.WriteLine("   <Row ss:Height=\"20\"><Cell ss:MergeAcross=\"" + (widths.Count-1) + "\" ss:StyleID=\"sFilterInfo\"><Data ss:Type=\"String\">🔍 " + SafeXml(detailFilter) + "</Data></Cell></Row>");
						sw.WriteLine("   <Row ss:Height=\"20\"><Cell ss:MergeAcross=\"" + (widths.Count-1) + "\" ss:StyleID=\"sSubtitle\"><Data ss:Type=\"String\">📅 Dicetak pada: " + DateTime.Now.ToString("dd MMMM yyyy HH:mm:ss") + " | 📈 Jumlah data: " + dt.Rows.Count + " transaksi</Data></Cell></Row>");
						sw.WriteLine("   <Row ss:Height=\"10\"></Row>");
            
						// Baris putih pertama dengan warna
						sw.WriteLine("   <Row ss:Height=\"8\">");
						for(int i = 0; i < widths.Count; i++)
							sw.WriteLine("    <Cell ss:StyleID=\"sDataBlue\"></Cell>");
						sw.WriteLine("   </Row>");
            
						// Baris putih kedua dengan warna  
						sw.WriteLine("   <Row ss:Height=\"8\">");
						for(int i = 0; i < widths.Count; i++)
							sw.WriteLine("    <Cell ss:StyleID=\"sDataBlue\"></Cell>");
						sw.WriteLine("   </Row>");

						// Menampilkan header tabel
						sw.WriteLine("   <Row>");
						
						for(int i = 0; i < allColumnNames.Length; i++) {
							if(dt.Columns.Contains(allColumnNames[i])) {
								string[] styleOptions = {"sHeaderBlue", "sHeaderGreen", "sHeaderOrange", "sHeaderRed", "sHeaderPurple"};
								string headerStyle = styleOptions[i % styleOptions.Length];
								sw.WriteLine("    <Cell ss:StyleID=\"" + headerStyle + "\"><Data ss:Type=\"String\">" + SafeXml(allHeaderTexts[i]) + "</Data></Cell>");
							}
						}
						sw.WriteLine("   </Row>");

						// Menampilkan data tabel
						double total = 0;
						int rowIndex = 0;
						int dataStartRow = 10; // Baris dimulai dari row 10 (setelah 2 baris kosong)
						int actualDataStartRow = dataStartRow; // Simpan nilai awal untuk rumus Excel
	            
						foreach(DataRow row in dt.Rows)
						{
							double nilai = 0;
							double nilaiKonversi = 0;
							double nilaiKurs = 0;
							try { nilai = Convert.ToDouble(row["nilai"]); total += nilai; } 
							catch {}
							try { nilaiKonversi = Convert.ToDouble(row["nilai_konversi"]); } 
							catch {}
							try { nilaiKurs = Convert.ToDouble(row["NILAI_KURS"]); } 
							catch {}

							sw.WriteLine("   <Row>");
							// Mengekspor data per baris
							for(int i = 0; i < allColumnNames.Length; i++) {
								if(dt.Columns.Contains(allColumnNames[i])) {
									object cellValue = row[allColumnNames[i]];
									string columnName = allColumnNames[i];
									
									if(columnName == "nilai" || columnName == "nilai_konversi" || columnName == "NILAI_KURS") {
										sw.WriteLine("    <Cell ss:StyleID=\"sNumber\"><Data ss:Type=\"Number\">" + cellValue.ToString() + "</Data></Cell>");
									} else {
										sw.WriteLine("    <Cell ss:StyleID=\"sDataBlue\"><Data ss:Type=\"String\">" + SafeXml(cellValue.ToString()) + "</Data></Cell>");
									}
								}
							}
							sw.WriteLine("   </Row>");
							rowIndex++;
							dataStartRow++;
						}

						// Baris putih di atas total dengan warna
						sw.WriteLine("   <Row ss:Height=\"12\">");
						for(int i = 0; i < widths.Count; i++)
							sw.WriteLine("    <Cell ss:StyleID=\"sDataBlue\"></Cell>");
						sw.WriteLine("   </Row>");

						// Total - Hanya satu total dengan rumus Excel di tengah
						sw.WriteLine("   <Row ss:Height=\"35\">");
						// Empty cells untuk spacing di kiri
						for(int i = 0; i < 3; i++)
							sw.WriteLine("    <Cell ss:StyleID=\"sTotal\"></Cell>");
						// Label total
						sw.WriteLine("    <Cell ss:MergeAcross=\"2\" ss:StyleID=\"sTotalLabel\"><Data ss:Type=\"String\">TOTAL KESELURUHAN</Data></Cell>");
						// Total dengan rumus Excel untuk menjumlah kolom NILAI (IDR)
						string sumFormula = "=SUM(R" + actualDataStartRow + "C4:R" + (actualDataStartRow + dt.Rows.Count - 1) + "C4)";
						
						// Tulis cell total dengan rumus Excel
						sw.WriteLine("    <Cell ss:StyleID=\"sTotalCentered\" ss:Formula=\"" + sumFormula + "\"><Data ss:Type=\"Number\">" + total + "</Data></Cell>");
						// Empty cells untuk spacing di kanan
						for(int i = 6; i < widths.Count; i++)
							sw.WriteLine("    <Cell ss:StyleID=\"sTotal\"></Cell>");
						sw.WriteLine("   </Row>");

						sw.WriteLine("  </Table>");
						sw.WriteLine(" </Worksheet>");
						sw.WriteLine("</Workbook>");

						// Menyimpan file Excel
						StreamWriter writer = new StreamWriter(save.FileName, false, Encoding.UTF8);
						writer.Write(sw.ToString());
						writer.Close();

						MessageBox.Show("Export Berhasil!");
					}
				}
				catch (Exception ex) 
				{ 
					MessageBox.Show("Gagal Export: " + ex.Message); 
				}
			}
		}
		private void dgTransaksi_Click(object sender, EventArgs e)
		{
			try 
			{
				int r = dgTransaksi.CurrentCell.RowNumber;
				DataTable dt = (DataTable)dgTransaksi.DataSource;
				if(dt != null && dt.Rows.Count > 0 && r < dt.Rows.Count)
				{
					// Mengambil nilai dengan aman
					if(dt.Columns.Contains("nilai") && dt.Rows[r]["nilai"] != DBNull.Value)
						txtDetailNilai.Text = Convert.ToInt64(dt.Rows[r]["nilai"]).ToString("N0", new CultureInfo("id-ID"));
					else
						txtDetailNilai.Text = "0";
					
					// Mengambil data keterangan
					if(dt.Columns.Contains("Ket_Header"))
						txtDetailKet.Text = dt.Rows[r]["Ket_Header"].ToString();
					else if(dt.Columns.Contains("keterangan"))
						txtDetailKet.Text = dt.Rows[r]["keterangan"].ToString();
					else
						txtDetailKet.Text = "";
				}
			} 
			catch (Exception ex) 
			{ 
				MessageBox.Show("Error: " + ex.Message); 
			}
		}

		private void MMTask_Load(object sender, EventArgs e) 
		{
			dtpMulai.CustomFormat = "dd MMM yyyy";
			dtpSelesai.CustomFormat = "dd MMM yyyy";
			cmbJenis.SelectedIndex = 0;
			cmbOperator.SelectedIndex = 2;
			SinkronisasiData();
		}

		private void btnCari_Click(object sender, EventArgs e) { SinkronisasiData(); }

		private void btnRefresh_Click(object sender, EventArgs e) 
		{
			txtCari.Text = ""; txtNominal.Text = ""; txtFilterKode.Text = ""; txtFilterKodeRekAju.Text = ""; cmbJenis.SelectedIndex = 0;
			cbPakaiTgl.Checked = false; SinkronisasiData();
		}

		#region Designer
		private void InitializeComponent()
		{
			this.panelSidebar = new System.Windows.Forms.Panel();
			this.labelCari = new System.Windows.Forms.Label();
			this.txtCari = new System.Windows.Forms.TextBox();
			this.lblNominal = new System.Windows.Forms.Label();
			this.cmbOperator = new System.Windows.Forms.ComboBox();
			this.txtNominal = new System.Windows.Forms.TextBox();
			this.labelKode = new System.Windows.Forms.Label();
			this.txtFilterKode = new System.Windows.Forms.TextBox();
			this.labelKodeRekAju = new System.Windows.Forms.Label();
			this.txtFilterKodeRekAju = new System.Windows.Forms.TextBox();
			this.lblJenis = new System.Windows.Forms.Label();
			this.cmbJenis = new System.Windows.Forms.ComboBox();
			this.cbPakaiTgl = new System.Windows.Forms.CheckBox();
			this.dtpMulai = new System.Windows.Forms.DateTimePicker();
			this.dtpSelesai = new System.Windows.Forms.DateTimePicker();
			this.btnCari = new System.Windows.Forms.Button();
			this.btnRefresh = new System.Windows.Forms.Button();
			this.btnExport = new System.Windows.Forms.Button();
			this.panelHeader = new System.Windows.Forms.Panel();
			this.labelInfoFilter = new System.Windows.Forms.Label();
			this.labelJudul = new System.Windows.Forms.Label();
			this.panelContent = new System.Windows.Forms.Panel();
			this.dgTransaksi = new System.Windows.Forms.DataGrid();
			this.lblTotalNilai = new System.Windows.Forms.Label();
			this.labelN = new System.Windows.Forms.Label();
			this.txtDetailNilai = new System.Windows.Forms.TextBox();
			this.labelK = new System.Windows.Forms.Label();
			this.txtDetailKet = new System.Windows.Forms.TextBox();
			this.lblStatus = new System.Windows.Forms.Label();
			this.panelSidebar.SuspendLayout();
			this.panelHeader.SuspendLayout();
			this.panelContent.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dgTransaksi)).BeginInit();
			this.SuspendLayout();
			// 
			// panelSidebar
			// 
			this.panelSidebar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(192)), ((System.Byte)(255)));
			this.panelSidebar.Controls.Add(this.labelCari);
			this.panelSidebar.Controls.Add(this.txtCari);
			this.panelSidebar.Controls.Add(this.lblNominal);
			this.panelSidebar.Controls.Add(this.cmbOperator);
			this.panelSidebar.Controls.Add(this.txtNominal);
			this.panelSidebar.Controls.Add(this.labelKode);
			this.panelSidebar.Controls.Add(this.txtFilterKode);
			this.panelSidebar.Controls.Add(this.labelKodeRekAju);
			this.panelSidebar.Controls.Add(this.txtFilterKodeRekAju);
			this.panelSidebar.Controls.Add(this.lblJenis);
			this.panelSidebar.Controls.Add(this.cmbJenis);
			this.panelSidebar.Controls.Add(this.cbPakaiTgl);
			this.panelSidebar.Controls.Add(this.dtpMulai);
			this.panelSidebar.Controls.Add(this.dtpSelesai);
			this.panelSidebar.Controls.Add(this.btnCari);
			this.panelSidebar.Controls.Add(this.btnRefresh);
			this.panelSidebar.Controls.Add(this.btnExport);
			this.panelSidebar.Dock = System.Windows.Forms.DockStyle.Left;
			this.panelSidebar.Location = new System.Drawing.Point(0, 0);
			this.panelSidebar.Name = "panelSidebar";
			this.panelSidebar.Size = new System.Drawing.Size(232, 1055);
			this.panelSidebar.TabIndex = 2;
			// 
			// labelCari
			// 
			this.labelCari.Location = new System.Drawing.Point(16, 16);
			this.labelCari.Name = "labelCari";
			this.labelCari.Size = new System.Drawing.Size(112, 16);
			this.labelCari.TabIndex = 0;
			this.labelCari.Text = "Keyword Umum:";
			// 
			// txtCari
			// 
			this.txtCari.Location = new System.Drawing.Point(16, 35);
			this.txtCari.Name = "txtCari";
			this.txtCari.Size = new System.Drawing.Size(155, 22);
			this.txtCari.TabIndex = 1;
			this.txtCari.Text = "";
			// 
			// lblNominal
			// 
			this.lblNominal.Location = new System.Drawing.Point(16, 65);
			this.lblNominal.Name = "lblNominal";
			this.lblNominal.Size = new System.Drawing.Size(100, 15);
			this.lblNominal.TabIndex = 2;
			this.lblNominal.Text = "Filter Nominal:";
			// 
			// cmbOperator
			// 
			this.cmbOperator.Items.AddRange(new object[] {
															 ">",
															 ">=",
															 "=",
															 "<",
															 "<="});
			this.cmbOperator.Location = new System.Drawing.Point(16, 85);
			this.cmbOperator.Name = "cmbOperator";
			this.cmbOperator.Size = new System.Drawing.Size(45, 24);
			this.cmbOperator.TabIndex = 3;
			// 
			// txtNominal
			// 
			this.txtNominal.Location = new System.Drawing.Point(65, 85);
			this.txtNominal.Name = "txtNominal";
			this.txtNominal.Size = new System.Drawing.Size(106, 22);
			this.txtNominal.TabIndex = 4;
			this.txtNominal.Text = "";
			this.txtNominal.Leave += new System.EventHandler(this.txtNominal_Leave);
			// 
			// labelKode
			// 
			this.labelKode.Location = new System.Drawing.Point(16, 115);
			this.labelKode.Name = "labelKode";
			this.labelKode.Size = new System.Drawing.Size(100, 13);
			this.labelKode.TabIndex = 5;
			this.labelKode.Text = "Kode Account:";
			// txtFilterKode
			this.txtFilterKode.Location = new System.Drawing.Point(16, 132);
			this.txtFilterKode.Name = "txtFilterKode";
			this.txtFilterKode.Size = new System.Drawing.Size(155, 22);
			this.txtFilterKode.TabIndex = 6;
			this.txtFilterKode.Text = "";
			// 
			// labelKodeRekAju
			// 
			this.labelKodeRekAju.Location = new System.Drawing.Point(16, 160);
			this.labelKodeRekAju.Name = "labelKodeRekAju";
			this.labelKodeRekAju.Size = new System.Drawing.Size(100, 13);
			this.labelKodeRekAju.TabIndex = 7;
			this.labelKodeRekAju.Text = "Kode Rek Aju:";
			// txtFilterKodeRekAju
			this.txtFilterKodeRekAju.Location = new System.Drawing.Point(16, 177);
			this.txtFilterKodeRekAju.Name = "txtFilterKodeRekAju";
			this.txtFilterKodeRekAju.Size = new System.Drawing.Size(155, 22);
			this.txtFilterKodeRekAju.TabIndex = 8;
			this.txtFilterKodeRekAju.Text = "";
			// 
			// lblJenis
			// 
			this.lblJenis.Location = new System.Drawing.Point(16, 205);
			this.lblJenis.Name = "lblJenis";
			this.lblJenis.Size = new System.Drawing.Size(100, 19);
			this.lblJenis.TabIndex = 9;
			this.lblJenis.Text = "Jenis:";
			// 
			// cmbJenis
			// 
			this.cmbJenis.Items.AddRange(new object[] {
														  "SEMUA",
														  "MASUK",
														  "KELUAR"});
			this.cmbJenis.Location = new System.Drawing.Point(16, 224);
			this.cmbJenis.Name = "cmbJenis";
			this.cmbJenis.Size = new System.Drawing.Size(155, 24);
			this.cmbJenis.TabIndex = 10;
			// 
			// cbPakaiTgl
			// 
			this.cbPakaiTgl.Location = new System.Drawing.Point(8, 264);
			this.cbPakaiTgl.Name = "cbPakaiTgl";
			this.cbPakaiTgl.Size = new System.Drawing.Size(136, 17);
			this.cbPakaiTgl.TabIndex = 9;
			this.cbPakaiTgl.Text = "Gunakan Tanggal";
			// 
			// dtpMulai
			// 
			this.dtpMulai.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpMulai.Location = new System.Drawing.Point(16, 288);
			this.dtpMulai.Name = "dtpMulai";
			this.dtpMulai.TabIndex = 10;
			// 
			// dtpSelesai
			// 
			this.dtpSelesai.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpSelesai.Location = new System.Drawing.Point(16, 320);
			this.dtpSelesai.Name = "dtpSelesai";
			this.dtpSelesai.TabIndex = 11;
			// 
			// btnCari
			// 
			this.btnCari.Location = new System.Drawing.Point(16, 368);
			this.btnCari.Name = "btnCari";
			this.btnCari.Size = new System.Drawing.Size(155, 40);
			this.btnCari.TabIndex = 12;
			this.btnCari.Text = "APPLY FILTER";
			this.btnCari.Click += new System.EventHandler(this.btnCari_Click);
			// 
			// btnRefresh
			// 
			this.btnRefresh.Location = new System.Drawing.Point(16, 416);
			this.btnRefresh.Name = "btnRefresh";
			this.btnRefresh.Size = new System.Drawing.Size(155, 40);
			this.btnRefresh.TabIndex = 13;
			this.btnRefresh.Text = "REFRESH";
			this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
			// btnExport
			this.btnExport.BackColor = System.Drawing.Color.SeaGreen;
			this.btnExport.ForeColor = System.Drawing.Color.White;
			this.btnExport.Location = new System.Drawing.Point(16, 740);
			this.btnExport.Name = "btnExport";
			this.btnExport.Size = new System.Drawing.Size(155, 50);
			this.btnExport.TabIndex = 14;
			this.btnExport.Text = "EXPORT EXCEL";
			this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
			// 
			// panelHeader
			// 
			this.panelHeader.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(41)), ((System.Byte)(128)), ((System.Byte)(185)));
			this.panelHeader.Controls.Add(this.labelInfoFilter);
			this.panelHeader.Controls.Add(this.labelJudul);
			this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
			this.panelHeader.Location = new System.Drawing.Point(232, 0);
			this.panelHeader.Name = "panelHeader";
			this.panelHeader.Size = new System.Drawing.Size(1510, 112);
			this.panelHeader.TabIndex = 1;
			// labelInfoFilter
			this.labelInfoFilter.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.labelInfoFilter.ForeColor = System.Drawing.Color.Yellow;
			this.labelInfoFilter.Location = new System.Drawing.Point(0, 82);
			this.labelInfoFilter.Name = "labelInfoFilter";
			this.labelInfoFilter.Size = new System.Drawing.Size(1510, 30);
			this.labelInfoFilter.TabIndex = 0;
			this.labelInfoFilter.Text = "Filter Aktif: SEMUA DATA";
			this.labelInfoFilter.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labelJudul
			// 
			this.labelJudul.Dock = System.Windows.Forms.DockStyle.Top;
			this.labelJudul.Font = new System.Drawing.Font("Tahoma", 18F, System.Drawing.FontStyle.Bold);
			this.labelJudul.ForeColor = System.Drawing.Color.White;
			this.labelJudul.Location = new System.Drawing.Point(0, 0);
			this.labelJudul.Name = "labelJudul";
			this.labelJudul.Size = new System.Drawing.Size(1510, 72);
			this.labelJudul.TabIndex = 1;
			this.labelJudul.Text = "DATA PENGAJUAN";
			this.labelJudul.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// panelContent
			// 
			this.panelContent.Controls.Add(this.dgTransaksi);
			this.panelContent.Controls.Add(this.lblTotalNilai);
			this.panelContent.Controls.Add(this.labelN);
			this.panelContent.Controls.Add(this.txtDetailNilai);
			this.panelContent.Controls.Add(this.labelK);
			this.panelContent.Controls.Add(this.txtDetailKet);
			this.panelContent.Controls.Add(this.lblStatus);
			this.panelContent.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panelContent.Location = new System.Drawing.Point(232, 112);
			this.panelContent.Name = "panelContent";
			this.panelContent.Size = new System.Drawing.Size(1510, 943);
			this.panelContent.TabIndex = 0;
			this.panelContent.Paint += new System.Windows.Forms.PaintEventHandler(this.panelContent_Paint);
			// 
			// dgTransaksi
			// 
			this.dgTransaksi.DataMember = "";
			this.dgTransaksi.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dgTransaksi.Location = new System.Drawing.Point(0, 0);
			this.dgTransaksi.Name = "dgTransaksi";
			this.dgTransaksi.Size = new System.Drawing.Size(1480, 580);
			this.dgTransaksi.TabIndex = 0;
			this.dgTransaksi.Click += new System.EventHandler(this.dgTransaksi_Click);
			// 
			// lblTotalNilai
			// 
			this.lblTotalNilai.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblTotalNilai.Location = new System.Drawing.Point(16, 590);
			this.lblTotalNilai.Name = "lblTotalNilai";
			this.lblTotalNilai.Size = new System.Drawing.Size(800, 25);
			this.lblTotalNilai.TabIndex = 1;
			// 
			// labelN
			// 
			this.labelN.Location = new System.Drawing.Point(16, 620);
			this.labelN.Name = "labelN";
			this.labelN.TabIndex = 2;
			this.labelN.Text = "Nilai Selected:";
			// 
			// txtDetailNilai
			// 
			this.txtDetailNilai.Location = new System.Drawing.Point(120, 620);
			this.txtDetailNilai.Name = "txtDetailNilai";
			this.txtDetailNilai.Size = new System.Drawing.Size(200, 22);
			this.txtDetailNilai.TabIndex = 3;
			this.txtDetailNilai.Text = "";
			this.txtDetailNilai.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// labelK
			// 
			this.labelK.Location = new System.Drawing.Point(16, 650);
			this.labelK.Name = "labelK";
			this.labelK.TabIndex = 4;
			this.labelK.Text = "Keterangan:";
			// 
			// txtDetailKet
			// 
			this.txtDetailKet.Location = new System.Drawing.Point(120, 650);
			this.txtDetailKet.Name = "txtDetailKet";
			this.txtDetailKet.Size = new System.Drawing.Size(800, 22);
			this.txtDetailKet.TabIndex = 5;
			this.txtDetailKet.Text = "";
			// 
			// lblStatus
			// 
			this.lblStatus.Location = new System.Drawing.Point(16, 680);
			this.lblStatus.Name = "lblStatus";
			this.lblStatus.Size = new System.Drawing.Size(624, 56);
			this.lblStatus.TabIndex = 6;
			// 
			// MMTask
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(1742, 1055);
			this.Controls.Add(this.panelContent);
			this.Controls.Add(this.panelHeader);
			this.Controls.Add(this.panelSidebar);
			this.Name = "MMTask";
			this.Text = "MM System - Data Pengajuan";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.MMTask_Load);
			this.panelSidebar.ResumeLayout(false);
			this.panelHeader.ResumeLayout(false);
			this.panelContent.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dgTransaksi)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		[STAThread] static void Main() { Application.Run(new MMTask()); }

		private void panelContent_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}
	}
}
