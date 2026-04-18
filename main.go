package main

import (
	"bufio"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// FakturSummary holds the header/footer data for each invoice PDF (Sheet 1).
type FakturSummary struct {
	NamaFile           string
	NomorFaktur        string
	NamaPenjual        string
	AlamatPenjual      string
	NPWPPenjual        string
	NamaPembeli        string
	AlamatPembeli      string
	NPWPPembeli        string
	NIKPembeli         string
	PasporPembeli      string
	IdenPembeli        string
	EmailPembeli       string
	HargaJualTotal     float64
	PotonganHargaTotal float64
	UangMuka           float64
	DPP                float64
	PPNTotal           float64
	PPnBMTotal         float64
	Tempat             string
	Tanggal            string
	Penandatangan      string
	Referensi          string
}

// BarangDetail holds the line-item detail for each good/service (Sheet 2).
type BarangDetail struct {
	NomorFaktur   string
	NamaBarang    string
	Harga         float64
	Qty           float64
	Kode          string
	Total         float64
	PotonganHarga float64
	TarifPPnBM    float64
	BesaranPPnBM  float64
}

// cleanNumber converts an Indonesian-formatted number string to float64.
// Dots are thousand separators, commas are decimal separators.
func cleanNumber(val string) float64 {
	if val == "" {
		return 0.0
	}
	clean := strings.ReplaceAll(val, ".", "")
	clean = strings.ReplaceAll(clean, ",", ".")
	f, err := strconv.ParseFloat(clean, 64)
	if err != nil {
		return 0.0
	}
	return f
}

// findPdftotext locates pdftotext.exe by checking:
// 1. Same directory as the running executable
// 2. Current working directory
// 3. System PATH
func findPdftotext() (string, error) {
	// Check next to executable
	exePath, err := os.Executable()
	if err == nil {
		candidate := filepath.Join(filepath.Dir(exePath), "pdftotext.exe")
		if _, err := os.Stat(candidate); err == nil {
			return candidate, nil
		}
	}

	// Check current working directory
	cwd, err := os.Getwd()
	if err == nil {
		candidate := filepath.Join(cwd, "pdftotext.exe")
		if _, err := os.Stat(candidate); err == nil {
			return candidate, nil
		}
	}

	// Check system PATH
	path, err := exec.LookPath("pdftotext")
	if err == nil {
		return path, nil
	}

	return "", fmt.Errorf("pdftotext.exe tidak ditemukan. Letakkan pdftotext.exe di folder yang sama dengan program ini")
}

// extractTextFromPDF uses pdftotext to extract text from a PDF, returns text and page count.
func extractTextFromPDF(pdftotextPath, filePath string) (string, int, error) {
	cmd := exec.Command(pdftotextPath, "-layout", filePath, "-")
	out, err := cmd.Output()
	if err != nil {
		return "", 0, fmt.Errorf("pdftotext error: %w", err)
	}

	text := string(out)

	// pdfplumber collapses multiple spaces frequently in its extract_text,
	// but mostly it doesn't - we will just clean the text a bit to help match python's behavior.
	// We will not completely collapse spaces because regex expects "\s+".

	// Count pages by counting form-feed characters (\f)
	pageCount := strings.Count(text, "\f")
	if pageCount == 0 && len(text) > 0 {
		pageCount = 1
	}

	return text, pageCount, nil
}

// promptInput reads a line from stdin with a prompt.
func promptInput(prompt string) string {
	fmt.Print(prompt)
	scanner := bufio.NewScanner(os.Stdin)
	scanner.Scan()
	return strings.TrimSpace(scanner.Text())
}

// extractMatch returns the first captured group from the regex, trimmed, or empty string.
func extractMatch(re *regexp.Regexp, text string) string {
	m := re.FindStringSubmatch(text)
	if m != nil && len(m) > 1 {
		return strings.TrimSpace(m[1])
	}
	return ""
}

// extractNumber returns the first captured group as float64 via cleanNumber.
func extractNumber(re *regexp.Regexp, text string) float64 {
	m := re.FindStringSubmatch(text)
	if m != nil && len(m) > 1 {
		return cleanNumber(m[1])
	}
	return 0.0
}

func main() {
	fmt.Println("Memulai proses ekstraksi faktur... Version 1.01 (Go)")

	// --- Locate pdftotext ---
	pdftotextPath, err := findPdftotext()
	if err != nil {
		fmt.Println(err)
		fmt.Println("\nDownload xpdf tools dari: https://www.xpdfreader.com/download.html")
		fmt.Println("Ekstrak dan letakkan pdftotext.exe di folder yang sama dengan program ini.")
		promptInput("\nTekan Enter untuk keluar...")
		os.Exit(1)
	}
	fmt.Printf("Menggunakan pdftotext: %s\n", pdftotextPath)

	// --- Konfigurasi Folder ---
	var folderPath string
	for {
		folderPath = promptInput("Masukkan path folder PDF: ")
		info, err := os.Stat(folderPath)
		if err == nil && info.IsDir() {
			break
		}
		fmt.Println("Folder tidak valid. Coba lagi.")
	}

	var outputExcel string
	for {
		outputExcel = promptInput("Masukkan path file output Excel (.xlsx): ")

		if !strings.HasSuffix(strings.ToLower(outputExcel), ".xlsx") {
			fmt.Println("File harus berekstensi .xlsx")
			continue
		}

		parentDir := filepath.Dir(outputExcel)
		if parentDir == "" {
			parentDir = "."
		}
		info, err := os.Stat(parentDir)
		if err != nil || !info.IsDir() {
			fmt.Println("Folder tujuan tidak ditemukan.")
			continue
		}

		break
	}

	// --- Scan PDF files ---
	entries, err := os.ReadDir(folderPath)
	if err != nil {
		fmt.Printf("Gagal membaca folder: %v\n", err)
		os.Exit(1)
	}

	var pdfFiles []string
	for _, e := range entries {
		if !e.IsDir() && strings.HasSuffix(strings.ToLower(e.Name()), ".pdf") {
			pdfFiles = append(pdfFiles, e.Name())
		}
	}

	totalFiles := len(pdfFiles)
	fmt.Printf("\nTotal file PDF ditemukan: %d\n", totalFiles)

	startTime := time.Now()

	var dataSummary []FakturSummary
	var dataDetail []BarangDetail

	// --- Regex patterns (compiled once) ---
	reFaktur := regexp.MustCompile(`Kode\s+dan\s+Nomor\s+Seri\s+Faktur\s+Pajak:\s+(\d+)`)
	reBlokPenjual := regexp.MustCompile(`(?s)Pengusaha Kena Pajak:(.*?)(?:Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:)`)
	reBlokPembeli := regexp.MustCompile(`(?s)Pembeli Barang Kena Pajak/Penerima Jasa Kena Pajak:(.*?)(?:Kode\s+Nama\s+Barang)`)

	reNamaPenjual := regexp.MustCompile(`(?s)Nama\s*:\s*(.*?)\s*Alamat\s*:`)
	reAlamatPenjual := regexp.MustCompile(`(?s)Alamat\s*:\s*(.*?)\s*NPWP\s*:`)
	reNPWPPenjual := regexp.MustCompile(`NPWP\s*:\s*([\d.\-]+)`)

	reNamaPembeli := regexp.MustCompile(`(?s)Nama\s*:\s*(.*?)\s*Alamat\s*:`)
	reAlamatPembeli := regexp.MustCompile(`(?s)Alamat\s*:\s*(.*?)\s*NPWP\s*:`)
	reNPWPPembeli := regexp.MustCompile(`(?s)NPWP\s*:\s*(.*?)\s*NIK\s*:`)
	reNIKPembeli := regexp.MustCompile(`(?s)NIK\s*:\s*(.*?)\s*Nomor\s*Paspor\s*:`)
	rePasporPembeli := regexp.MustCompile(`(?s)Nomor\s*Paspor\s*:\s*(.*?)\s*Identitas\s*Lain\s*:`)
	reIdenPembeli := regexp.MustCompile(`(?s)Identitas\s*Lain\s*:\s*(.*?)\s*Email\s*:`)
	reEmailPembeli := regexp.MustCompile(`Email\s*:\s*([^\n\r]+)`)

	reHargaJual := regexp.MustCompile(`Harga Jual / Penggantian / Uang Muka / Termin\s+([\d.,]+)`)
	rePotongan := regexp.MustCompile(`Dikurangi Potongan Harga\s+([\d.,]+)`)
	reDPP := regexp.MustCompile(`Dasar Pengenaan Pajak\s+([\d.,]+)`)
	reUangMuka := regexp.MustCompile(`Dikurangi Uang Muka yang telah diterima\s+([\d.,]+)`)
	rePPN := regexp.MustCompile(`Jumlah PPN \(Pajak Pertambahan Nilai\)\s+([\d.,]+)`)
	rePPnBM := regexp.MustCompile(`Jumlah PPnBM \(Pajak Penjualan atas Barang Mewah\)\s+([\d.,]+)`)

	reBlokSplit := regexp.MustCompile(`(?s)PPnBM.*?\n`)
	reHargaLine := regexp.MustCompile(`Rp\s*([\d.,]+)\s*x\s*([\d.,]+)`)
	reKode := regexp.MustCompile(`\b\d\s+(\d{6})\b`)
	rePotonganBarang := regexp.MustCompile(`Potongan\s*Harga\s*=\s*([^\n\r]+)`)
	rePPnBMBarang := regexp.MustCompile(`PPnBM\s*\((.*?)\)\s*=\s*(Rp\s*[\d.,]+)`)
	reRef := regexp.MustCompile(`Referensi:\s*(.*?)\)`)

	for idx, filename := range pdfFiles {
		filePath := filepath.Join(folderPath, filename)

		fmt.Printf("\n[%d/%d] Memproses: %s\n", idx+1, totalFiles, filename)

		fullText, numPages, err := extractTextFromPDF(pdftotextPath, filePath)
		if err != nil {
			fmt.Printf("Gagal memproses file %s: %v\n", filename, err)
			continue
		}

		fmt.Printf("   -> Jumlah halaman: %d\n", numPages)

		// Create a copy of fullText where multiple spaces are collapsed to a single space ONLY FOR BLOK MATCHING
		// Because pdfplumber joins words with single space, so pdftotext -layout's big gaps can break simple matchers
		// But wait, the Python `Kode Harga Jual` or `Jasa (Rp)` regex relies on the exact sequence.
		// Let's just use `Kode\s+Nama\s+Barang` for Python's `Kode Harga Jual` since pdftotext inserts spaces.
		// Let's also fix `Jasa (Rp)` matching. In pdftotext -layout it says `Jasa` on one line and `(Rp)` on another line sometimes.

		nomorFaktur := "Tidak Ditemukan"
		if m := reFaktur.FindStringSubmatch(fullText); m != nil {
			nomorFaktur = m[1]
		}

		penjualText := ""
		if m := reBlokPenjual.FindStringSubmatch(fullText); m != nil {
			penjualText = m[1]
		}
		pembeliText := ""
		if m := reBlokPembeli.FindStringSubmatch(fullText); m != nil {
			pembeliText = m[1]
		}

		// Penjual
		namaPenjual := extractMatch(reNamaPenjual, penjualText)
		alamatPenjual := extractMatch(reAlamatPenjual, penjualText)
		npwpPenjual := extractMatch(reNPWPPenjual, penjualText)

		// Pembeli
		namaPembeli := extractMatch(reNamaPembeli, pembeliText)
		alamatPembeli := extractMatch(reAlamatPembeli, pembeliText)
		npwpPembeli := extractMatch(reNPWPPembeli, pembeliText)
		nikPembeli := extractMatch(reNIKPembeli, pembeliText)
		pasporPembeli := extractMatch(rePasporPembeli, pembeliText)
		idenPembeli := extractMatch(reIdenPembeli, pembeliText)
		emailPembeli := extractMatch(reEmailPembeli, pembeliText)

		// --- Merge overlapping address block ---
		// Karena pdftotext -layout dapat memecah baris Alamat yang panjang hingga sejajar dengan baris Nama:
		// Nama  : PT AAAA            RT 01, RW 02
		// Alamat: JL BBB
		mergeExtraAddress := func(nama, alamat string) (string, string) {
			reSpaces := regexp.MustCompile(` {3,}`)
			parts := reSpaces.Split(nama, 2)
			if len(parts) > 1 {
				nama = strings.TrimSpace(parts[0])
				extraAddress := strings.TrimSpace(parts[1])
				
				// Gabungkan extraAddress ke dalam alamat
				// Biasanya nyelip di antara baris 1 dan baris 2 dari blok alamat
				alamat = strings.ReplaceAll(alamat, "\r\n", "\n")
				alamatLines := strings.SplitN(alamat, "\n", 2)
				if len(alamatLines) > 1 {
					alamat = strings.TrimSpace(alamatLines[0]) + " " + extraAddress + "\n" + strings.TrimSpace(alamatLines[1])
				} else {
					alamat = strings.TrimSpace(alamat) + " " + extraAddress
				}
			} else {
				nama = strings.TrimSpace(nama)
				// Clean newlines
				alamat = strings.ReplaceAll(alamat, "\r\n", "\n")
			}
			return nama, alamat
		}

		namaPenjual, alamatPenjual = mergeExtraAddress(namaPenjual, alamatPenjual)
		namaPembeli, alamatPembeli = mergeExtraAddress(namaPembeli, alamatPembeli)
		hargaJualTotal := extractNumber(reHargaJual, fullText)
		potonganTotal := extractNumber(rePotongan, fullText)
		dpp := extractNumber(reDPP, fullText)
		uangMuka := extractNumber(reUangMuka, fullText)
		ppnTotal := extractNumber(rePPN, fullText)
		ppnbmTotal := extractNumber(rePPnBM, fullText)

		// Tandatangan - matching python exactly
		var lines []string
		for _, l := range strings.Split(fullText, "\n") {
			trimmed := strings.TrimSpace(l)
			if trimmed != "" {
				lines = append(lines, trimmed)
			}
		}

		var tempat, tanggal, penandatangan, ref string
		anchorText := "Ditandatangani secara elektronik"
		anchorIdx := -1
		for i, s := range lines {
			if strings.Contains(s, anchorText) {
				anchorIdx = i
				break
			}
		}

		if anchorIdx >= 0 {
			if anchorIdx > 0 {
				lineAtas := lines[anchorIdx-1]
				if idx := strings.Index(lineAtas, ","); idx >= 0 {
					tempat = strings.TrimSpace(lineAtas[:idx])
					tanggal = strings.TrimSpace(lineAtas[idx+1:])
				} else {
					tempat = lineAtas
				}
			}
			if anchorIdx+1 < len(lines) {
				penandatangan = lines[anchorIdx+1]
			}
			for _, s := range lines[anchorIdx:] {
				if strings.Contains(s, "Referensi:") {
					if m := reRef.FindStringSubmatch(s); m != nil {
						ref = strings.TrimSpace(m[1])
					}
					break
				}
			}
		} // Cleanup format like python did on `name.group(1).strip()` vs `name.group(1)`
		dataSummary = append(dataSummary, FakturSummary{
			NamaFile:           filename,
			NomorFaktur:        nomorFaktur,
			NamaPenjual:        namaPenjual,
			AlamatPenjual:      alamatPenjual,
			NPWPPenjual:        npwpPenjual,
			NamaPembeli:        namaPembeli,
			AlamatPembeli:      alamatPembeli,
			NPWPPembeli:        strings.TrimSpace(npwpPembeli),
			NIKPembeli:         strings.TrimSpace(nikPembeli),
			PasporPembeli:      strings.TrimSpace(pasporPembeli),
			IdenPembeli:        strings.TrimSpace(idenPembeli),
			EmailPembeli:       strings.TrimSpace(emailPembeli),
			HargaJualTotal:     hargaJualTotal,
			PotonganHargaTotal: potonganTotal,
			UangMuka:           uangMuka,
			DPP:                dpp,
			PPNTotal:           ppnTotal,
			PPnBMTotal:         ppnbmTotal,
			Tempat:             tempat,
			Tanggal:            tanggal,
			Penandatangan:      penandatangan,
			Referensi:          ref,
		})

		// --- Ekstraksi Tabel Barang (Sheet 2) ---
		// Allow numbers after Jasa (e.g., 1.207.207,20)
		reBlokBarang := regexp.MustCompile(`(?s)\(Rp\)\s*\n+\s*Jasa[\d.,\s]*\n+(.*?)(?:Harga Jual /)`)
		if m := reBlokBarang.FindStringSubmatch(fullText); m != nil {
			blokBarang := m[1]
			blocks := reBlokSplit.Split(blokBarang, -1)

			for _, block := range blocks {
				block = strings.TrimSpace(block)
				if block == "" {
					continue
				}

				var bLines []string
				for _, l := range strings.Split(block, "\n") {
					trimmed := strings.TrimSpace(l)
					if trimmed != "" {
						bLines = append(bLines, trimmed)
					}
				}

				hargaIdx := -1
				for i, l := range bLines {
					if strings.Contains(l, "Rp") && strings.Contains(l, "x") {
						hargaIdx = i
						break
					}
				}

				if hargaIdx >= 0 {
					// Use bLines[:hargaIdx] but drop lines that are purely numbers/commas/dots
					var nameLines []string
					rePureNum := regexp.MustCompile(`^[\d.,]+$`)
					for _, l := range bLines[:hargaIdx] {
						if !rePureNum.MatchString(l) {
							nameLines = append(nameLines, l)
						}
					}

					namaBarang := strings.Join(nameLines, " ")
					// Remove trailing numbers (priced formats like 45.045,04)
					reTrailingNum := regexp.MustCompile(`\s+[\d.]+,[\d]{2}$`)
					namaBarang = reTrailingNum.ReplaceAllString(namaBarang, "")
					// Remove multiple spaces inside name to match python's clean string
					reSpaces := regexp.MustCompile(`\s+`)
					namaBarang = reSpaces.ReplaceAllString(namaBarang, " ")
					namaBarang = strings.TrimSpace(namaBarang)

					var harga, qty string
					if hm := reHargaLine.FindStringSubmatch(bLines[hargaIdx]); hm != nil {
						harga = hm[1]
						qty = hm[2]
					}
					total := cleanNumber(qty) * cleanNumber(harga)

					var kode string
					if km := reKode.FindStringSubmatch(block); km != nil {
						kode = strings.TrimSpace(km[1])
					}

					var potonganHarga string
					if pm := rePotonganBarang.FindStringSubmatch(block); pm != nil {
						potonganHarga = strings.TrimSpace(pm[1])
						potonganHarga = strings.ReplaceAll(potonganHarga, "Rp", "")
						potonganHarga = strings.TrimSpace(potonganHarga)
					}

					var tarifPPnBM, besaranPPnBM string
					if bm := rePPnBMBarang.FindStringSubmatch(block); bm != nil {
						tarifPPnBM = bm[1]
						besaranPPnBM = bm[2]

						besaranPPnBM = strings.ReplaceAll(besaranPPnBM, "Rp", "")
						besaranPPnBM = strings.TrimSpace(besaranPPnBM)
					}

					dataDetail = append(dataDetail, BarangDetail{
						NomorFaktur:   nomorFaktur,
						NamaBarang:    namaBarang,
						Harga:         cleanNumber(harga),
						Qty:           cleanNumber(qty),
						Kode:          kode,
						Total:         total,
						PotonganHarga: cleanNumber(potonganHarga),
						TarifPPnBM:    cleanNumber(tarifPPnBM),
						BesaranPPnBM:  cleanNumber(besaranPPnBM),
					})
				}
			}
		}
	}

	elapsed := time.Since(startTime)
	fmt.Printf("\nWaktu proses: %.2f detik\n", elapsed.Seconds())

	// --- Export ke Excel ---
	xlsx := excelize.NewFile()
	defer xlsx.Close()

	// Sheet 1: Summary Faktur
	summarySheet := "Summary Faktur"
	xlsx.SetSheetName("Sheet1", summarySheet)

	summaryHeaders := []string{
		"nama_file", "nomor_faktur",
		"nama_penjual", "alamat_penjual", "npwp_penjual",
		"nama_pembeli", "alamat_pembeli", "npwp_pembeli",
		"nik_pembeli", "paspor_pembeli", "iden_pembeli", "email_pembeli",
		"harga_jual_total", "potongan_harga_total", "uang_muka",
		"dpp", "ppn_total", "ppnbm_total",
		"tempat", "tanggal", "penandatangan", "referensi",
	}
	for col, h := range summaryHeaders {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		xlsx.SetCellValue(summarySheet, cell, h)
	}

	for row, s := range dataSummary {
		vals := []interface{}{
			s.NamaFile, s.NomorFaktur,
			s.NamaPenjual, s.AlamatPenjual, s.NPWPPenjual,
			s.NamaPembeli, s.AlamatPembeli, s.NPWPPembeli,
			s.NIKPembeli, s.PasporPembeli, s.IdenPembeli, s.EmailPembeli,
			s.HargaJualTotal, s.PotonganHargaTotal, s.UangMuka,
			s.DPP, s.PPNTotal, s.PPnBMTotal,
			s.Tempat, s.Tanggal, s.Penandatangan, s.Referensi,
		}
		for col, v := range vals {
			cell, _ := excelize.CoordinatesToCellName(col+1, row+2)
			xlsx.SetCellValue(summarySheet, cell, v)
		}
	}

	// Sheet 2: Detail Barang
	detailSheet := "Detail Barang"
	xlsx.NewSheet(detailSheet)

	detailHeaders := []string{
		"nomor_faktur", "nama_barang", "harga", "qty",
		"kode", "total", "potongan_harga",
		"tarif_ppnbm", "besaran_ppnbm",
	}
	for col, h := range detailHeaders {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		xlsx.SetCellValue(detailSheet, cell, h)
	}

	for row, d := range dataDetail {
		vals := []interface{}{
			d.NomorFaktur, d.NamaBarang, d.Harga, d.Qty,
			d.Kode, d.Total, d.PotonganHarga,
			d.TarifPPnBM, d.BesaranPPnBM,
		}
		for col, v := range vals {
			cell, _ := excelize.CoordinatesToCellName(col+1, row+2)
			xlsx.SetCellValue(detailSheet, cell, v)
		}
	}

	// Format kolom Qty (Kolom D) di Sheet Detail Barang sebagai desimal (0.00)
	if styleID, err := xlsx.NewStyle(&excelize.Style{NumFmt: 2}); err == nil {
		xlsx.SetColStyle(detailSheet, "D", styleID)
	}

	if err := xlsx.SaveAs(outputExcel); err != nil {
		fmt.Printf("Gagal menyimpan Excel: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("Proses selesai. Data disimpan di %s\n", outputExcel)
	promptInput("\nTekan Enter untuk keluar... Version 1.01 (Go)")
}
