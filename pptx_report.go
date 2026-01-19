package main

import (
	"bytes"
	"encoding/csv"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strings"
	"time"

	chart "github.com/wcharczuk/go-chart/v2"
)

type pptxManifest struct {
	Title  string          `json:"title"`
	Slides []pptxSlideSpec `json:"slides"`
}

type pptxSlideSpec struct {
	Title     string `json:"title"`
	ImagePath string `json:"image"`
}

func maybeGeneratePPTX(csvPath, pptxFlag string, periodStart time.Time) error {
	pptxFlag = strings.TrimSpace(pptxFlag)
	if pptxFlag == "" {
		return nil
	}

	pptxPath := pptxFlag
	if strings.EqualFold(pptxFlag, "auto") {
		pptxPath = defaultPPTXName(periodStart)
	}

	absPPTX := mustAbs(pptxPath)
	pngDir := strings.TrimSuffix(absPPTX, filepath.Ext(absPPTX)) + "_png"
	if err := os.MkdirAll(pngDir, 0o755); err != nil {
		return fmt.Errorf("create png dir: %w", err)
	}

	slides, err := buildPiePNGsFromCSV(csvPath, pngDir)
	if err != nil {
		return err
	}
	if len(slides) == 0 {
		return errors.New("no slides generated (no data?)")
	}

	manifest := pptxManifest{
		Title:  fmt.Sprintf("Relatório %04d-%02d", periodStart.Year(), int(periodStart.Month())),
		Slides: slides,
	}
	manifestPath := filepath.Join(pngDir, "manifest.json")
	b, err := json.MarshalIndent(manifest, "", "  ")
	if err != nil {
		return fmt.Errorf("marshal manifest: %w", err)
	}
	if err := os.WriteFile(manifestPath, b, 0o644); err != nil {
		return fmt.Errorf("write manifest: %w", err)
	}

	if err := runPythonPPTXBuilder(manifestPath, absPPTX); err != nil {
		return err
	}

	fmt.Printf("OK: PPTX gerado em %s (PNGs em %s)\n", absPPTX, pngDir)
	return nil
}

func defaultPPTXName(periodStart time.Time) string {
	return fmt.Sprintf("relatorio_%04d_%02d.pptx", periodStart.Year(), int(periodStart.Month()))
}

func buildPiePNGsFromCSV(csvPath, pngDir string) ([]pptxSlideSpec, error) {
	f, err := os.Open(csvPath)
	if err != nil {
		return nil, fmt.Errorf("open csv: %w", err)
	}
	defer f.Close()

	r := csv.NewReader(f)
	r.Comma = ';'
	r.FieldsPerRecord = -1

	headerRow, err := r.Read()
	if err != nil {
		return nil, fmt.Errorf("read header: %w", err)
	}
	if len(headerRow) > 0 {
		headerRow[0] = strings.TrimPrefix(headerRow[0], "\ufeff") // handle UTF-8 BOM
	}

	// Expected layout from our exporter:
	// 0 ANDAR
	// 1 Paciente
	// 2..21 questao1..questao20
	// 22 Data - Criação
	// 23 Cadastrador
	if len(headerRow) < 24 {
		return nil, fmt.Errorf("csv has %d columns; expected >= 24", len(headerRow))
	}

	questionCols := questionColumns(headerRow)

	counts := make([]map[string]int, len(questionCols))
	for i := range counts {
		counts[i] = map[string]int{}
	}

	for {
		row, err := r.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("read csv: %w", err)
		}
		for i, qc := range questionCols {
			if qc.Index >= len(row) {
				continue
			}
			v := strings.TrimSpace(row[qc.Index])
			if v == "" {
				continue
			}
			v = replaceValue(v) // normalize numeric codes when present
			counts[i][v]++
		}
	}

	slides := make([]pptxSlideSpec, 0, len(questionCols))
	for i, qc := range questionCols {
		values := counts[i]
		if len(values) == 0 {
			continue
		}
		pngBytes, err := renderPiePNG(values)
		if err != nil {
			return nil, fmt.Errorf("render pie for %s: %w", qc.Title, err)
		}
		imgName := fmt.Sprintf("q%02d.png", qc.Number)
		imgPath := filepath.Join(pngDir, imgName)
		if err := os.WriteFile(imgPath, pngBytes, 0o644); err != nil {
			return nil, fmt.Errorf("write png %s: %w", imgName, err)
		}
		slides = append(slides, pptxSlideSpec{Title: qc.Title, ImagePath: imgPath})
	}

	return slides, nil
}

type questionCol struct {
	Number int
	Index  int
	Title  string
}

func questionColumns(headerRow []string) []questionCol {
	// Exclude:
	// - questao16 => CSV index 17
	// - questao20 => CSV index 21
	// Mapping: questaoN => index = 1 + N (because 0 andar, 1 paciente)
	cols := make([]questionCol, 0, 18)
	for n := 1; n <= 20; n++ {
		idx := 1 + n
		if n == 16 || n == 20 {
			continue
		}
		cols = append(cols, questionCol{Number: n, Index: idx, Title: strings.TrimSpace(headerRow[idx])})
	}
	return cols
}

func renderPiePNG(counts map[string]int) ([]byte, error) {
	total := 0
	for _, c := range counts {
		total += c
	}
	if total <= 0 {
		return nil, errors.New("empty counts")
	}

	type kv struct {
		K string
		V int
	}
	items := make([]kv, 0, len(counts))
	for k, v := range counts {
		items = append(items, kv{K: k, V: v})
	}
	sort.Slice(items, func(i, j int) bool {
		if items[i].V != items[j].V {
			return items[i].V > items[j].V
		}
		return items[i].K < items[j].K
	})

	values := make([]chart.Value, 0, len(items))
	for _, it := range items {
		pct := (float64(it.V) / float64(total)) * 100
		label := fmt.Sprintf("%s (%d - %.1f%%)", it.K, it.V, pct)
		values = append(values, chart.Value{Value: float64(it.V), Label: label})
	}

	pie := chart.PieChart{
		Width:  1024,
		Height: 768,
		Values: values,
	}

	var buf bytes.Buffer
	if err := pie.Render(chart.PNG, &buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func runPythonPPTXBuilder(manifestPath, pptxOutPath string) error {
	py := pythonExecutablePath()
	script := "pptx_builder.py"
	if _, err := os.Stat(script); err != nil {
		// If the user runs the binary from another working dir, try alongside the binary.
		if exe, exeErr := os.Executable(); exeErr == nil {
			candidate := filepath.Join(filepath.Dir(exe), "pptx_builder.py")
			if _, statErr := os.Stat(candidate); statErr == nil {
				script = candidate
			}
		}
	}

	cmd := exec.Command(py, script, "--manifest", manifestPath, "--out", pptxOutPath)
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	if err := cmd.Run(); err != nil {
		return fmt.Errorf("python pptx builder failed: %w", err)
	}
	return nil
}

func pythonExecutablePath() string {
	// Prefer project venv if present.
	venvRel := filepath.Join(".venv", "Scripts", "python.exe")
	if _, err := os.Stat(venvRel); err == nil {
		return venvRel
	}
	if exe, err := os.Executable(); err == nil {
		venvNextToExe := filepath.Join(filepath.Dir(exe), ".venv", "Scripts", "python.exe")
		if _, statErr := os.Stat(venvNextToExe); statErr == nil {
			return venvNextToExe
		}
	}
	return "python"
}
