package main

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
)

func readJSONRecords(path string) ([]map[string]string, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}

	var records []map[string]string
	err = json.Unmarshal(data, &records)
	if err != nil {
		return nil, err
	}

	return records, nil
}

func readCSVRecords(path string) ([]map[string]string, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	reader := csv.NewReader(file)
	rows, err := reader.ReadAll()
	if err != nil {
		return nil, err
	}

	if len(rows) < 2 {
		return nil, fmt.Errorf("CSV must have at least a header row and one data row")
	}

	headers := rows[0]
	var records []map[string]string

	for i := 1; i < len(rows); i++ {
		if len(rows[i]) != len(headers) {
			continue // Skip malformed rows
		}

		record := make(map[string]string)
		for j, header := range headers {
			record[header] = rows[i][j]
		}
		records = append(records, record)
	}

	return records, nil
}
