package llm

type Provider struct {
	Name      string
	UserKey   string
	SecretKey string
	uri       string
}

type LlmJob struct {
	SystemPrompt string
	UserPrompt   string
	Response     string
}

type ProviderHelper interface {
	// SetConfig(config Provider)
	// GetConfig() Provider
	Call(job LlmJob)
}
