# Claude Code with Git Worktrees: Complete Guide

## Overview

Git worktrees enable you to run multiple instances of Claude Code simultaneously on the same project by creating isolated working directories that share a single Git repository. This guide covers everything you need to know about using worktrees with Claude Code for parallel development.

## Why Use Worktrees with Claude Code?

### Key Benefits
- **Parallel Development**: Run multiple Claude sessions simultaneously, each working on different features
- **Isolation**: Prevents Claude instances from interfering with each other's changes
- **Efficiency**: Shares Git history and repository data (unlike multiple clones)
- **Speed**: Theoretically doubles productivity by allowing concurrent development

### When to Use
- Working on multiple independent features simultaneously
- Testing different approaches to the same problem
- Long-running features that benefit from parallel development
- Complex projects requiring isolated experimentation

## Basic Setup and Commands

### Prerequisites
- Git version 2.5 or higher
- Claude Code installed and configured
- A Git repository with at least one commit

### Step-by-Step Setup

1. **Navigate to your main project**
   ```bash
   cd /path/to/your/project
   ```

2. **Create a new worktree**
   ```bash
   git worktree add ../project-worktrees/feature-name -b feature/branch-name
   ```
   This creates:
   - A new directory at `../project-worktrees/feature-name`
   - A new branch called `feature/branch-name`
   - An isolated working directory linked to your repository

3. **Open separate terminal tabs/windows**
   - Terminal 1: Stay in main project directory
   - Terminal 2: Navigate to worktree directory
   ```bash
   cd ../project-worktrees/feature-name
   ```

4. **Start Claude Code sessions**
   - In Terminal 1: `claude` (for main branch work)
   - In Terminal 2: `claude` (for feature branch work)

5. **Work independently**
   - Each Claude instance can modify files without affecting the other
   - Changes are isolated to their respective branches

### Essential Commands

```bash
# List all worktrees
git worktree list

# Create worktree from existing branch
git worktree add ../worktrees/bugfix bugfix/issue-123

# Create worktree with new branch
git worktree add -b feature/new-feature ../worktrees/new-feature

# Remove a worktree (after committing changes)
git worktree remove ../worktrees/feature-name

# Force remove a worktree (with uncommitted changes)
git worktree remove --force ../worktrees/feature-name

# Prune stale worktree entries
git worktree prune
```

## Advanced Automation

### Custom Worktree Manager Function

Create a bash function `w` for streamlined worktree management:

```bash
# Add to ~/.bashrc or ~/.zshrc
function w() {
    local repo=$1
    local worktree=$2
    local command=${@:3}
    
    # Base directory for all worktrees
    local WORKTREE_BASE="$HOME/projects/worktrees"
    
    # If no arguments, list worktrees
    if [[ -z "$repo" ]]; then
        ls -la "$WORKTREE_BASE" 2>/dev/null
        return
    fi
    
    # Create worktree path
    local worktree_path="$WORKTREE_BASE/$repo/$worktree"
    
    # If worktree doesn't exist, create it
    if [[ ! -d "$worktree_path" ]]; then
        echo "Creating worktree: $worktree_path"
        mkdir -p "$(dirname "$worktree_path")"
        
        # Navigate to main repo
        cd "$HOME/projects/$repo"
        
        # Create worktree with branch
        git worktree add "$worktree_path" -b "$USER/$worktree"
    fi
    
    # If command provided, execute in worktree context
    if [[ -n "$command" ]]; then
        cd "$worktree_path" && $command
    else
        # Otherwise, just navigate to worktree
        cd "$worktree_path"
    fi
}

# Add auto-completion (for zsh)
compdef _w w
function _w() {
    # Auto-complete repository and worktree names
    # Implementation depends on your shell
}
```

### Usage Examples

```bash
# Create and enter new worktree
w excel-explorer feature-screenshots

# Start Claude in specific worktree
w excel-explorer feature-screenshots claude

# Run commands without changing directory
w excel-explorer feature-screenshots git status
w excel-explorer feature-screenshots npm install
w excel-explorer feature-screenshots python main.py

# Quick commit from anywhere
w excel-explorer feature-screenshots git commit -m "feat: add screenshot capture"
```

## MCP Templates for Different Workflows

### Setting Up MCP Profiles

Create template MCP configurations for different types of work:

1. **Create template directory**
   ```bash
   mkdir ~/.claude/mcp-templates
   ```

2. **Create specialized templates**

   **Data Analysis Template** (`~/.claude/mcp-templates/data-analysis.json`):
   ```json
   {
     "mcpServers": {
       "filesystem": {
         "command": "npx",
         "args": ["-y", "@modelcontextprotocol/server-filesystem", "/path/to/data"]
       },
       "memory-bank": {
         "command": "npx",
         "args": ["-y", "@allpepper/memory-bank-mcp"]
       }
     }
   }
   ```

   **Web Development Template** (`~/.claude/mcp-templates/web-dev.json`):
   ```json
   {
     "mcpServers": {
       "puppeteer": {
         "command": "npx",
         "args": ["-y", "@modelcontextprotocol/server-puppeteer"]
       },
       "browser-tools": {
         "command": "npx",
         "args": ["-y", "@agentdeskai/browser-tools-mcp"]
       }
     }
   }
   ```

3. **Automate template copying**
   ```bash
   function setup-mcp() {
       local template=$1
       cp ~/.claude/mcp-templates/$template.json ./.mcp.json
       echo "MCP profile '$template' applied to current directory"
   }
   ```

4. **Use with worktrees**
   ```bash
   w myproject feature-api
   setup-mcp api-development
   claude
   ```

## Best Practices

### Organization
- **Naming Convention**: Use descriptive names like `feature/user-auth` or `bugfix/api-error`
- **Directory Structure**: Keep worktrees in a dedicated parent directory (e.g., `~/worktrees/`)
- **Branch Prefixes**: Use your username or initials as prefix for branches

### Development Workflow
1. **One Feature Per Worktree**: Keep each worktree focused on a single feature
2. **Regular Commits**: Commit frequently in each worktree to preserve work
3. **Clean Up**: Remove worktrees after merging to avoid clutter
4. **Documentation**: Keep a README in each worktree describing its purpose

### Performance Considerations
- **Token Usage**: Running multiple Claude sessions consumes more API tokens
- **Context Switching**: Mental overhead of managing multiple contexts
- **System Resources**: Each worktree uses disk space for working files
- **Setup Time**: Initial setup can be time-consuming for complex projects

### Environment Setup
For each new worktree, remember to:
- **Node.js Projects**: Run `npm install` or `yarn install`
- **Python Projects**: Set up virtual environment and install dependencies
- **Database Projects**: Configure connection strings if needed
- **Environment Variables**: Copy or link `.env` files as appropriate

## Common Issues and Solutions

### Issue: Worktree Already Exists
```bash
# Error: 'path/to/worktree' already exists
git worktree remove path/to/worktree
git worktree prune
# Then recreate
```

### Issue: Branch Already Checked Out
```bash
# Error: 'branch-name' is already checked out
# Find which worktree has it
git worktree list
# Switch that worktree to different branch or remove it
```

### Issue: Uncommitted Changes
```bash
# Before removing worktree, either:
# 1. Commit changes
cd path/to/worktree
git add .
git commit -m "WIP: saving work"

# 2. Or force remove (loses changes!)
git worktree remove --force path/to/worktree
```

### Issue: MCP Servers Not Available
```bash
# Each worktree needs its own .mcp.json
# Copy from main project or use template
cp /main/project/.mcp.json ./
```

## Alternative Approaches

### GitButler Integration
GitButler offers an alternative where multiple Claude Code instances can work in the same directory:
- Automatically creates branches for each session
- Uses Claude Code hooks for file tracking
- Assigns work to appropriate branches automatically
- Creates commits when chats complete

### Crystal Desktop App
Crystal is an Electron application for managing multiple Claude Code instances:
- GUI for worktree management
- Visual inspection of parallel sessions
- Built-in testing capabilities
- Simplified workflow management

## Example: Complete Parallel Development Session

```bash
# Main project: Refactoring core module
cd ~/projects/excel-explorer
claude  # Start session 1

# New terminal: Adding new feature
git worktree add ../excel-explorer-worktrees/screenshot-feature -b feature/screenshots
cd ../excel-explorer-worktrees/screenshot-feature
pip install -r requirements.txt  # Set up environment
claude  # Start session 2

# New terminal: Fixing bug
git worktree add ../excel-explorer-worktrees/bugfix -b bugfix/data-validation
cd ../excel-explorer-worktrees/bugfix
pip install -r requirements.txt
claude  # Start session 3

# After work is complete
cd ~/projects/excel-explorer
git worktree list  # Review active worktrees

# Merge and clean up
git merge feature/screenshots
git merge bugfix/data-validation
git worktree remove ../excel-explorer-worktrees/screenshot-feature
git worktree remove ../excel-explorer-worktrees/bugfix
```

## Conclusion

Git worktrees with Claude Code enable powerful parallel development workflows. While there's initial setup overhead, the ability to run multiple isolated Claude sessions simultaneously can significantly boost productivity on complex projects. Choose this approach when working on multiple independent features or when you need to test different solutions in parallel.

Remember to:
- Keep worktrees organized and named clearly
- Clean up after merging
- Monitor token usage
- Set up environments properly in each worktree
- Use automation tools to reduce friction

This approach works best for experienced developers comfortable with Git and command-line workflows who need to maximize development velocity on complex projects.