package helpers

import (
	"fmt"
	"strconv"
	"strings"
)

// A1ToGrid converts A1 notation (e.g., "B5") to 0-indexed grid coordinates
func A1ToGrid(cell string) (col int, row int, err error) {
	i := 0
	for i < len(cell) && cell[i] >= 'A' && cell[i] <= 'Z' {
		col = col*26 + int(cell[i]-'A'+1)
		i++
	}

	if i < len(cell) {
		rowNum, err := strconv.Atoi(cell[i:])
		if err != nil {
			return 0, 0, fmt.Errorf("invalid cell reference: %s", cell)
		}
		row = rowNum
	}

	col--
	row--

	return col, row, nil
}

// ParseColumns parses column letter(s) or range (e.g., "A", "B:D", "AA") into 0-indexed start/end column indices.
func ParseColumns(s string) (start, end int, err error) {
	s = strings.ToUpper(strings.TrimSpace(s))
	parts := strings.Split(s, ":")

	start, err = columnLetterToIndex(parts[0])
	if err != nil {
		return 0, 0, err
	}

	if len(parts) > 1 {
		end, err = columnLetterToIndex(parts[1])
		if err != nil {
			return 0, 0, err
		}
	} else {
		end = start
	}

	return start, end, nil
}

// columnLetterToIndex converts column letters (e.g., "A"=0, "Z"=25, "AA"=26) to 0-indexed column index.
func columnLetterToIndex(col string) (int, error) {
	if col == "" {
		return 0, fmt.Errorf("empty column reference")
	}
	idx := 0
	for _, c := range col {
		if c < 'A' || c > 'Z' {
			return 0, fmt.Errorf("invalid column letter: %s", col)
		}
		idx = idx*26 + int(c-'A'+1)
	}
	return idx - 1, nil
}

// ParseRange parses A1 notation range (e.g., "A1:B10") into 0-indexed grid coordinates
func ParseRange(rangeA1 string) (startCol, startRow, endCol, endRow int, err error) {
	parts := strings.Split(rangeA1, ":")

	startCol, startRow, err = A1ToGrid(parts[0])
	if err != nil {
		return 0, 0, 0, 0, err
	}

	if len(parts) > 1 {
		endCol, endRow, err = A1ToGrid(parts[1])
		if err != nil {
			return 0, 0, 0, 0, err
		}
	} else {
		endCol, endRow = startCol, startRow
	}

	return startCol, startRow, endCol, endRow, nil
}
