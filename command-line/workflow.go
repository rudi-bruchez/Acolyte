package main

import llm "pachadata.com/ghostwriter/v2/llm"

type Runner struct {
	provider llm.ProviderHelper
	job      llm.LlmJob
}

func NewRunner(provider llm.ProviderHelper, job llm.LlmJob) Runner {
	return Runner{
		provider: provider,
		job:      job,
	}
}

func (runner Runner) runJob() {
	// ...
	runner.provider.Call(runner.job)
}
