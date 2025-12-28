package helpers

const (
	FormatTypeCurrency = "CURRENCY"
	FormatTypeDate     = "DATE"
	FormatTypeNumber   = "NUMBER"
	FormatTypePercent  = "PERCENT"
	FormatTypeTime     = "TIME"
)

// Default format patterns for each format type
var defaultFormatPatterns = map[string]string{
	FormatTypeNumber:   "#,##0.00",
	FormatTypeCurrency: "$#,##0.00",
	FormatTypeDate:     "yyyy-mm-dd",
	FormatTypePercent:  "0.00%",
	FormatTypeTime:     "hh:mm:ss",
}

// GetDefaultFormatPattern returns the default pattern for a given format type
func GetDefaultFormatPattern(formatType string) string {
	if pattern, ok := defaultFormatPatterns[formatType]; ok {
		return pattern
	}
	return ""
}
