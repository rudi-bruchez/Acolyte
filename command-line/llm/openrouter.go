package llm

import (
	"context"
	"fmt"

	openai "github.com/sashabaranov/go-openai"
)

// implements the ProviderHelper interface
type OpenRouterProvider struct {
	Provider
}

// func (p *OpenAiProvider) SetConfig(config Provider) {
// 	p.Name = config.Name
// 	p.UserKey = config.UserKey
// 	p.SecretKey = config.SecretKey
// 	p.Uri = "https://api.openai.com/v1"
// }

// func (p *OpenAiProvider) GetConfig() Provider {
// 	return p.Provider
// }

func (p *OpenRouterProvider) Call(job LlmJob) {
	p.uri = "https://api.openrouter.com/v1"

	// Call the LLM
	client := openai.NewClient("your token")
	resp, err := client.CreateChatCompletion(
		context.Background(),
		openai.ChatCompletionRequest{
			Model: openai.GPT3Dot5Turbo,
			Messages: []openai.ChatCompletionMessage{
				{
					Role:    openai.ChatMessageRoleUser,
					Content: "Hello!",
				},
			},
		},
	)

	if err != nil {
		fmt.Printf("ChatCompletion error: %v\n", err)
		return
	}

	fmt.Println(resp.Choices[0].Message.Content)
}
