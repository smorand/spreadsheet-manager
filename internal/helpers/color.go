package helpers

import (
	"strconv"
	"strings"

	"google.golang.org/api/sheets/v4"
)

const (
	HexColorLength = 6
	RGBMaxValue    = 255.0
)

// ParseColor converts hex color string (e.g., "#ff0000") to Google Sheets Color object
// Color values are normalized to 0.0-1.0 range
func ParseColor(hexColor string) *sheets.Color {
	hexColor = strings.TrimPrefix(hexColor, "#")
	if len(hexColor) != HexColorLength {
		return nil
	}

	r, err1 := strconv.ParseInt(hexColor[0:2], 16, 64)
	g, err2 := strconv.ParseInt(hexColor[2:4], 16, 64)
	b, err3 := strconv.ParseInt(hexColor[4:6], 16, 64)

	if err1 != nil || err2 != nil || err3 != nil {
		return nil
	}

	return &sheets.Color{
		Red:   float64(r) / RGBMaxValue,
		Green: float64(g) / RGBMaxValue,
		Blue:  float64(b) / RGBMaxValue,
	}
}
