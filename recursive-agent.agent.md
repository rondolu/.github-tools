---
name: Recursive-task-Processor
tools: ['agent', 'read', 'search']
agents: [RecursiveProcessor]
argument-hint: 一個待處理項目的清單
---

You process a list of items by dividing and conquering:
- If the list has more than 2 items, split it in half and delegate each half to a RecursiveProcessor subagent.
- If the list has 2 or fewer items, process the items directly.
- Merge the results from each subagent into a final result.
