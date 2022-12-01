package main

import (
	"fmt"
	"log"
	"os"
)

const DefaultConfigPath = "./config.json"

func main() {
	if len(os.Args) < 2 {
		fmt.Fprintln(os.Stderr, "Usage: timefiller [path to xlsx file]")
		os.Exit(-1)
	}

	originalFilePath := os.Args[1]

	config, configErr := LoadConfig(DefaultConfigPath)
	if configErr != nil {
		log.Fatalf("Error loading config: %s\n", configErr)
	}

	_, err := FillFile(originalFilePath, &config)
	if err != nil {
		log.Fatalf("Error filling XLSX file '%s': %s", originalFilePath, err)
	}

	log.Println("Done.")
}
