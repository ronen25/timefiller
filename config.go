package main

import (
	"encoding/json"
	"fmt"
	"os"
)

type Config struct {
	EmployeeName string `json:"employee_name"`
}

func DefaultConfig() Config {
	return Config{
		EmployeeName: "name here",
	}
}

func LoadConfig(path string) (Config, error) {
	// If a config doesn't exist, try to generate it
	if _, err := os.Stat(path); err != nil {
		file, err := os.Create(path)
		if err != nil {
			return Config{}, err
		}
		defer file.Close()

		data, err := json.MarshalIndent(DefaultConfig(), "", "  ")
		if err != nil {
			return Config{}, err
		}

		_, err = file.WriteString(string(data))
		if err != nil {
			return Config{}, fmt.Errorf("failed to generate default config: %s", err)
		}
	}

	data, err := os.ReadFile(path)
	if err != nil {
		return Config{}, err
	}

	var config Config
	jsonErr := json.Unmarshal(data, &config)
	if jsonErr != nil {
		return Config{}, jsonErr
	}

	return config, nil
}
