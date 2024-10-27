package main

import (
	"context"
	"flag"
	"fmt"
	"io"
	"log"
	"os"

	"github.com/joho/godotenv"
	"github.com/tmc/langchaingo/llms"
	"github.com/tmc/langchaingo/llms/openai"
	"pachadata.com/acolyte/v2/llm"
)

type Config struct {
	ApiKey string
}

func main() {
	file := flag.String("conf", "", "Configuration file")
	flag.Parse()

	if !IsFlagPassed("conf") {
		fmt.Println("Usage: acolyte -conf <configuration file>")
		return
	}

	var env map[string]string
	env, err := godotenv.Read()

	// Récupération de la valeur de la variable d'environnement API_KEY
	conf := Config{
		ApiKey: env["API_KEY"],
	}

	// Affichage de la valeur de la variable d'environnement API_KEY
	fmt.Printf("API Key: %s\n", conf.ApiKey)

	// Ouverture du fichier JSON
	configFile, err := os.Open(*file)
	if err != nil {
		log.Fatalf("Erreur lors de l'ouverture du fichier: %v", err)
	}
	defer configFile.Close()

	// Lecture du contenu du fichier
	bytes, err := io.ReadAll(configFile)
	if err != nil {
		log.Fatalf("Erreur lors de la lecture du fichier: %v", err)
	}

	// Décodage du JSON dans la struct Config
	var work Workload
	err = work.Load(bytes)
	if err != nil {
		log.Fatalf("Erreur lors du chargement de la configuration: %v", err)
	}

	// log.Fatal(http.ListenAndServe(":"+port, nil))

	// Création d'un nouveau Runner
	provider := &llm.OpenAiProvider{}

	provider.SecretKey = conf.ApiKey

	job := llm.LlmJob{
		SystemPrompt: work.SystemPrompt,
		UserPrompt:   work.MainPrompt,
	}

	runner := NewRunner(provider, job)
	runner.runJob()

	langchain(work.MainPrompt)
}

func langchain(prompt string) {
	llm, err := openai.New()
	if err != nil {
		log.Fatal(err)
	}
	ctx := context.Background()
	completion, err := llm.Call(ctx, prompt,
		llms.WithTemperature(0.8),
		llms.WithStopWords([]string{"Armstrong"}),
	)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println(completion)
}
