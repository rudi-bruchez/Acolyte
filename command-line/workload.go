package main

import (
	"encoding/json"
	"fmt"
	"log"
)

// Config représente la structure du fichier de configuration JSON
type Workload struct {
	Title        string   `json:"title"`
	Type         string   `json:"type"`
	OutputFormat string   `json:"output_format"`
	Description  string   `json:"description"`
	Keywords     []string `json:"keywords"`
	Authors      []string `json:"authors"`
	Date         string   `json:"date"`
	Lang         string   `json:"lang"`
	SystemPrompt string   `json:"system_prompt"`
	MainPrompt   string   `json:"main_prompt"`
	Structure    struct{} `json:"structure"`
	Content      []struct {
		Type    string `json:"type"`
		Title   string `json:"title"`
		Content []struct {
			Type    string `json:"type"`
			Content string `json:"content"`
		} `json:"content"`
	} `json:"content"`
	Steps []struct {
		Type   string `json:"type"`
		Title  string `json:"title"`
		Prompt string `json:"prompt"`
		Source string `json:"source"`
	} `json:"steps"`
}

func (work Workload) Load(bytes []byte) error {
	err := json.Unmarshal(bytes, &work)
	if err != nil {
		log.Fatalf("Erreur lors du décodage JSON: %v", err)
		return err
	}

	// Affichage des informations lues
	fmt.Printf("Titre: %s\n", work.Title)
	fmt.Printf("Type: %s\n", work.Type)
	fmt.Printf("Format de sortie: %s\n", work.OutputFormat)
	fmt.Printf("Description: %s\n", work.Description)
	fmt.Printf("Mots-clés: %v\n", work.Keywords)
	fmt.Printf("Auteurs: %v\n", work.Authors)
	fmt.Printf("Date: %s\n", work.Date)
	fmt.Printf("Langue: %s\n", work.Lang)
	fmt.Printf("System Prompt: %s\n", work.SystemPrompt)
	fmt.Printf("Main Prompt: %s\n", work.MainPrompt)

	return nil
}
