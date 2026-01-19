package main

import (
	"context"
	"database/sql"
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"

	_ "github.com/go-sql-driver/mysql"
	"github.com/joho/godotenv"
)

// Contract:
// - Inputs:
//   - MySQL DSN via --dsn or MYSQL_DSN env
//   - Period via --start/--end (RFC3339) OR --month/--year (month closed)
//   - Output CSV via --out (default: ./export.csv)
// - Output:
//   - CSV file with header row + rows from query
//
// Notes:
// - Uses a half-open interval [start, end) to avoid 23:59:59 problems.

const query = `
SELECT
    l.num_andar,
    p.nome_paciente,
    eq.questao1,
    eq.questao2,
    eq.questao3,
    eq.questao4,
    eq.questao5,
    eq.questao6,
    eq.questao7,
    eq.questao8,
    eq.questao9,
    eq.questao10,
    eq.questao11,
    eq.questao12,
    eq.questao13,
    eq.questao14,
    eq.questao15,
    eq.questao16,
    eq.questao17,
    eq.questao18,
    eq.questao19,
    eq.questao20,
    eq.created,
    eq.cadastrador
FROM adms_experiencia_questoes AS eq
LEFT JOIN adms_leitos AS l
    ON eq.adms_leito_id = l.id
LEFT JOIN adms_paciente AS p
    ON eq.adms_paciente_id = p.id
WHERE eq.created >= ?
  AND eq.created <  ?
ORDER BY eq.created ASC;
`

var header = []string{
	"ANDAR",
	"Paciente",
	"ATENDIMENTO DE RECEPÇÃO/ORIENTAÇÃO",
	"ATENDIMENTO MÉDICO",
	"ATENDIMENTO DE ENFERMAGEM",
	"ATENDIMENTO REGULAÇÃO",
	"ATENDIMENTO EQUIPE MULTI(PSICOLOGIA / SERVIÇO SOCIAL / NUTRIÇÃO)",
	"ATENDIMENTO DE EXAMES DIAGNÓSTICOS",
	"ATENDIMENTO TELEFÔNICO",
	"LIMPEZA DA UNIDADE",
	"INSTALAÇÕES",
	"TEMPO DE ESPERA DO ATENDIMENTO",
	"Recomendaria esse hospital para seus amigos e familiares?",
	"Teve confirmado em algum momento do seu atendimento seu nome e data de nascimento?",
	"Recebeu informações sobre a continuidade de seu tratamento?",
	"Foi adequadamente orientado quanto a forma de utilização de suas medicações?",
	"SEU PROBLEMA DE SAÚDE FOI RESOLVIDO OU CONTROLADO NO HOSPITAL DIA ?",
	"CASO NÃO, EXPLIQUE O PORQUÊ :",
	"SE ALIMENTA AO MÍNIMO COM 5 PORÇÕES DE FRUTAS, VERDURAS E LEGUMES DIARIAMENTE?",
	"Você foi atendido com gentileza e empatia? Sentiu nossos colaboradores motivados?",
	"Tempo de acesso e de retorno na especialidade",
	"O QUE IMPORTA PARA VOCÊ EM NOSSO SERVIÇO:",
	"Data - Criação",
	"Cadastrador",
}

func main() {
	// Carrega variáveis do arquivo .env (se existir) para evitar passar tudo via cmd.
	// Flags continuam tendo precedência, porque são lidas depois.
	_ = godotenv.Load()

	var (
		dsn       = flag.String("dsn", "", "MySQL DSN. If empty, uses MYSQL_DSN env. Example: user:pass@tcp(host:3306)/db?parseTime=true&charset=utf8mb4")
		out       = flag.String("out", "", "Output CSV path (optional). If empty, auto-generates name based on month/year.")
		pptxOut   = flag.String("pptx", "", "Optional PowerPoint (.pptx) output path. If set to 'auto', generates relatorio_YYYY_MM.pptx and a PNG folder next to it.")
		pptxFrom  = flag.String("pptx-from", "", "Generate PPTX from an existing CSV file and exit (skips DB query). Requires --pptx or --pptx=auto.")
		start     = flag.String("start", "", "Start datetime (RFC3339). Example: 2025-12-01T00:00:00-03:00")
		end       = flag.String("end", "", "End datetime (RFC3339, exclusive). Example: 2026-01-01T00:00:00-03:00")
		month     = flag.Int("month", 0, "Month number 1-12 (alternative to --start/--end)")
		year      = flag.Int("year", 0, "Year (alternative to --start/--end)")
		repl      = flag.Bool("replace", false, "Replace numeric codes in questao1..questao20 (like the VBA macro: 1..7 -> text)")
		bom       = flag.Bool("bom", true, "Write UTF-8 BOM at start of CSV (recommended for Excel)")
		dedupe    = flag.Bool("dedupe", true, "Remove consecutive duplicate rows when Paciente and Data - Criação indicate duplicates")
		dedupeSec = flag.Int("dedupe-sec", 60, "Dedup tolerance in seconds for consecutive rows with same Paciente (default 60). Use 0 for strict timestamp equality")
	)
	flag.Parse()

	if strings.TrimSpace(*pptxFrom) != "" {
		if strings.TrimSpace(*pptxOut) == "" {
			log.Fatal("when using --pptx-from, you must set --pptx or --pptx=auto")
		}
		if err := maybeGeneratePPTX(*pptxFrom, *pptxOut, time.Now()); err != nil {
			log.Fatalf("pptx: %v", err)
		}
		return
	}

	dsnVal, err := resolveDSN(*dsn)
	if err != nil {
		log.Fatal(err)
	}

	periodStart, periodEnd, err := resolvePeriod(*start, *end, *month, *year)
	if err != nil {
		log.Fatalf("invalid period: %v", err)
	}

	// Se --out não foi informado, gera automaticamente um nome (mês/ano do período).
	outPath := *out
	if strings.TrimSpace(outPath) == "" {
		outPath = defaultOutName(periodStart)
	}

	if err := os.MkdirAll(filepath.Dir(mustAbs(outPath)), 0o755); err != nil && filepath.Dir(outPath) != "." {
		log.Fatalf("create output dir: %v", err)
	}

	db, err := sql.Open("mysql", ensureParseTime(dsnVal))
	if err != nil {
		log.Fatalf("open db: %v", err)
	}
	defer db.Close()

	ctx, cancel := context.WithTimeout(context.Background(), 2*time.Minute)
	defer cancel()

	if err := db.PingContext(ctx); err != nil {
		log.Fatalf("ping db: %v", err)
	}

	rows, err := db.QueryContext(ctx, query, periodStart, periodEnd)
	if err != nil {
		log.Fatalf("query: %v", err)
	}
	defer rows.Close()

	f, err := os.Create(outPath)
	if err != nil {
		log.Fatalf("create csv: %v", err)
	}
	defer f.Close()

	if *bom {
		// Excel costuma interpretar CSV como ANSI/Windows-1252 sem BOM.
		// Escrevendo BOM UTF-8 (EF BB BF), ele detecta UTF-8 e mantém acentos (ã, ç, é...).
		if _, err := f.Write([]byte{0xEF, 0xBB, 0xBF}); err != nil {
			log.Fatalf("write BOM: %v", err)
		}
	}

	w := csv.NewWriter(f)
	w.Comma = ';' // padrão comum pt-BR/Excel. Se quiser vírgula, troque para ','

	if err := w.Write(header); err != nil {
		log.Fatalf("write header: %v", err)
	}

	count := 0
	skipped := 0
	var prevPaciente string
	var prevCreated string
	var prevCreatedTime time.Time
	var hasPrev bool
	for rows.Next() {
		record, err := scanRowToStrings(rows)
		if err != nil {
			log.Fatalf("scan row: %v", err)
		}
		if *repl {
			record = applyReplacements(record)
		}

		if *dedupe {
			// Layout do record esperado:
			// 0 ANDAR
			// 1 Paciente
			// 2..21 questões
			// 22 Data - Criação (YYYY-MM-DD HH:MM:SS)
			// 23 Cadastrador
			if len(record) >= 24 {
				paciente := strings.TrimSpace(record[1])
				created := strings.TrimSpace(record[22])

				if hasPrev && paciente != "" && created != "" && paciente == prevPaciente {
					// strict compare
					if *dedupeSec <= 0 {
						if created == prevCreated {
							skipped++
							continue
						}
					} else {
						// tolerant compare: parse time and consider duplicates if within N seconds
						curT, okCur := parseCreated(created)
						prevT, okPrev := prevCreatedTime, !prevCreatedTime.IsZero()
						if okCur && okPrev {
							d := curT.Sub(prevT)
							if d < 0 {
								d = -d
							}
							if d <= time.Duration(*dedupeSec)*time.Second {
								skipped++
								continue
							}
						} else {
							// fallback: if we can't parse, fall back to strict string compare
							if created == prevCreated {
								skipped++
								continue
							}
						}
					}
				}

				prevPaciente, prevCreated = paciente, created
				prevCreatedTime, _ = parseCreated(created)
				hasPrev = true
			}
		}

		if err := w.Write(record); err != nil {
			log.Fatalf("write row: %v", err)
		}
		count++
	}
	if err := rows.Err(); err != nil {
		log.Fatalf("rows: %v", err)
	}

	w.Flush()
	if err := w.Error(); err != nil {
		log.Fatalf("flush csv: %v", err)
	}

	if *dedupe {
		fmt.Printf("OK: %d linhas exportadas (removidas %d duplicadas consecutivas) para %s (%s -> %s)\n", count, skipped, outPath, periodStart.Format(time.RFC3339), periodEnd.Format(time.RFC3339))
		if err := maybeGeneratePPTX(outPath, *pptxOut, periodStart); err != nil {
			log.Fatalf("pptx: %v", err)
		}
		return
	}
	fmt.Printf("OK: %d linhas exportadas para %s (%s -> %s)\n", count, outPath, periodStart.Format(time.RFC3339), periodEnd.Format(time.RFC3339))
	if err := maybeGeneratePPTX(outPath, *pptxOut, periodStart); err != nil {
		log.Fatalf("pptx: %v", err)
	}
}

func parseCreated(s string) (time.Time, bool) {
	// Expected: YYYY-MM-DD HH:MM:SS
	t, err := time.ParseInLocation("2006-01-02 15:04:05", strings.TrimSpace(s), time.Local)
	if err != nil {
		return time.Time{}, false
	}
	return t, true
}

func defaultOutName(periodStart time.Time) string {
	// Nome baseado no mês/ano do período selecionado.
	// Ex.: relatorio_2026_01.csv
	return fmt.Sprintf("relatorio_%04d_%02d.csv", periodStart.Year(), int(periodStart.Month()))
}

func applyReplacements(record []string) []string {
	// record layout:
	// 0 num_andar
	// 1 nome_paciente
	// 2..21 questao1..questao20
	// 22 created
	// 23 cadastrador
	// A macro VBA aplicava em C:U (20 colunas) -> aqui equivale a questao1..questao20.
	if len(record) < 24 {
		return record
	}
	for i := 2; i <= 21; i++ {
		// questao16 e questao20 são strings: não aplicar replace.
		// Indices no record:
		// - questao1  => 2
		// - questao16 => 17
		// - questao20 => 21
		if i == 17 || i == 21 {
			continue
		}
		record[i] = replaceValue(record[i])
	}
	return record
}

func replaceValue(v string) string {
	// Normaliza espaços e aceita valores como "1", "1.0", " 1 ".
	s := strings.TrimSpace(v)
	if s == "" {
		return v
	}
	// Se vier "1.0" do banco/export, pega a parte inteira.
	if strings.Contains(s, ".") {
		parts := strings.SplitN(s, ".", 2)
		if len(parts) > 0 {
			s = parts[0]
		}
	}

	switch s {
	case "1":
		return "Ruim"
	case "2":
		return "Boa"
	case "3":
		return "Regular"
	case "4":
		return "Excelente"
	case "5":
		return "Não utilizei"
	case "6":
		return "Sim"
	case "7":
		return "Não"
	default:
		return v
	}
}

func scanRowToStrings(rows *sql.Rows) ([]string, error) {
	// num_andar pode ser NULL dependendo do join. nome_paciente idem.
	var (
		numAndar     sql.NullString
		nomePaciente sql.NullString
		questoes     [20]sql.NullString
		created      sql.NullTime
		cadastrador  sql.NullString
	)

	dests := make([]any, 0, 2+20+2)
	dests = append(dests, &numAndar, &nomePaciente)
	for i := 0; i < 20; i++ {
		dests = append(dests, &questoes[i])
	}
	dests = append(dests, &created, &cadastrador)

	if err := rows.Scan(dests...); err != nil {
		return nil, err
	}

	rec := make([]string, 0, len(header))
	rec = append(rec, nullToString(numAndar))
	rec = append(rec, nullToString(nomePaciente))
	for i := 0; i < 20; i++ {
		rec = append(rec, nullToString(questoes[i]))
	}
	if created.Valid {
		rec = append(rec, created.Time.Format("2006-01-02 15:04:05"))
	} else {
		rec = append(rec, "")
	}
	rec = append(rec, mapCadastrador(nullToString(cadastrador)))
	return rec, nil
}

func mapCadastrador(v string) string {
	// Ajuste solicitado: valor 5 deve aparecer como nome completo.
	s := strings.TrimSpace(v)
	if s == "" {
		return v
	}
	// Se vier como "5.0" por algum motivo, considera como 5
	if strings.Contains(s, ".") {
		parts := strings.SplitN(s, ".", 2)
		if len(parts) > 0 {
			s = parts[0]
		}
	}
	if s == "5" {
		return "Edna das Graças Prates Cruz"
	}
	return v
}

func resolvePeriod(start, end string, month, year int) (time.Time, time.Time, error) {
	if start != "" || end != "" {
		if start == "" || end == "" {
			return time.Time{}, time.Time{}, errors.New("use both --start and --end")
		}
		s, err := time.Parse(time.RFC3339, start)
		if err != nil {
			return time.Time{}, time.Time{}, fmt.Errorf("parse --start: %w", err)
		}
		e, err := time.Parse(time.RFC3339, end)
		if err != nil {
			return time.Time{}, time.Time{}, fmt.Errorf("parse --end: %w", err)
		}
		if !e.After(s) {
			return time.Time{}, time.Time{}, errors.New("--end must be after --start")
		}
		return s, e, nil
	}

	if month == 0 && year == 0 {
		// default: month closed = previous month in local time
		now := time.Now()
		firstThisMonth := time.Date(now.Year(), now.Month(), 1, 0, 0, 0, 0, now.Location())
		firstPrevMonth := firstThisMonth.AddDate(0, -1, 0)
		return firstPrevMonth, firstThisMonth, nil
	}
	if month < 1 || month > 12 {
		return time.Time{}, time.Time{}, errors.New("--month must be 1..12")
	}
	if year < 2000 || year > 2100 {
		return time.Time{}, time.Time{}, errors.New("--year must be between 2000 and 2100")
	}

	loc := time.Local
	s := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, loc)
	e := s.AddDate(0, 1, 0)
	return s, e, nil
}

func nullToString(ns sql.NullString) string {
	if ns.Valid {
		return ns.String
	}
	return ""
}

func firstNonEmpty(a, b string) string {
	if a != "" {
		return a
	}
	return b
}

func resolveDSN(flagDSN string) (string, error) {
	// Precedência:
	// 1) --dsn
	// 2) MYSQL_DSN
	// 3) Monta DSN a partir de MYSQL_HOST/PORT/DB/USER/PASS
	if flagDSN != "" {
		return flagDSN, nil
	}
	if envDSN := os.Getenv("MYSQL_DSN"); envDSN != "" {
		return envDSN, nil
	}

	host := firstNonEmpty(os.Getenv("MYSQL_HOST"), "127.0.0.1")
	port := firstNonEmpty(os.Getenv("MYSQL_PORT"), "3306")
	db := os.Getenv("MYSQL_DB")
	user := os.Getenv("MYSQL_USER")
	pass := os.Getenv("MYSQL_PASS")

	if db == "" || user == "" {
		return "", errors.New("missing DSN: set MYSQL_DSN or set MYSQL_DB and MYSQL_USER (and optionally MYSQL_PASS, MYSQL_HOST, MYSQL_PORT)")
	}

	// Importante: o DSN do go-sql-driver/mysql NÃO é uma URL.
	// Se fizermos QueryEscape na senha (ex.: '@' -> '%40'), a senha muda e o MySQL nega acesso.
	// Então montamos o DSN com user/pass exatamente como você digitou no .env.
	dsn := fmt.Sprintf("%s:%s@tcp(%s:%s)/%s?charset=utf8mb4&parseTime=true", user, pass, host, port, db)
	return dsn, nil
}

func mustAbs(p string) string {
	abs, err := filepath.Abs(p)
	if err != nil {
		return p
	}
	return abs
}

func ensureParseTime(dsn string) string {
	// go-sql-driver/mysql needs parseTime=true to scan DATETIME/TIMESTAMP into time.Time reliably.
	if hasQueryParam(dsn, "parseTime") {
		return dsn
	}
	sep := "?"
	if contains(dsn, "?") {
		sep = "&"
	}
	return dsn + sep + "parseTime=true"
}

func hasQueryParam(dsn, key string) bool {
	// minimal check; DSN format is not a full URL.
	return contains(dsn, key+"=")
}

func contains(s, sub string) bool {
	return len(sub) == 0 || (len(s) >= len(sub) && (index(s, sub) >= 0))
}

func index(s, sub string) int {
	// strings.Index, but avoiding another import for a tiny file.
	for i := 0; i+len(sub) <= len(s); i++ {
		if s[i:i+len(sub)] == sub {
			return i
		}
	}
	return -1
}
