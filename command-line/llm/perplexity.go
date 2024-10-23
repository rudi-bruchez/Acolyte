package llm

import (
	"context"
	"fmt"

	openai "github.com/sashabaranov/go-openai"
)

type PerplexityProvider struct {
	Provider
}

// func (p *PerplexityProvider) SetConfig() {
// 	p.Name = "Perplexity"
// 	p.UserKey = "user_key"
// 	p.SecretKey = "secret"
// }

// func (p *PerplexityProvider) GetConfig() {
// 	return
// }

func (p *PerplexityProvider) Call(job LlmJob) {
	p.uri = "https://api.perplexity.com/v1"

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
