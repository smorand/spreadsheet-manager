package helpers

import (
	"encoding/json"
	"fmt"
)

// PrintJSON marshals a value to indented JSON and prints it to stdout
func PrintJSON(v interface{}) error {
	data, err := json.MarshalIndent(v, "", "  ")
	if err != nil {
		return err
	}
	fmt.Println(string(data))
	return nil
}
