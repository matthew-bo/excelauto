# SheetSense Design Document

## 📘 Product Design Document (PRD)

### 1. Overview

**Purpose:**
SheetSense is a native Excel Add-in that empowers users to interact with spreadsheets through natural language. It enables:

* AI explanations of formulas, data structures, and workbook logic.
* AI-powered creation, editing, and formatting of spreadsheet content from user prompts.
* Structured, previewed, and reversible spreadsheet actions similar to Rocode.

**Target Users:**

* Financial analysts
* Business professionals
* Students and Excel learners
* Operations teams managing models

**Why Now:**
Spreadsheets are essential but increasingly complex. SheetSense bridges the gap between business logic and Excel mechanics, enabling users to say "Create a dashboard" or "Explain this formula," and have it executed instantly.

---

### 2. Core Features

| Category      | Feature                                    | Description                                                            |
| ------------- | ------------------------------------------ | ---------------------------------------------------------------------- |
| Understanding | Formula Explanation                        | AI explains selected formulas                                          |
|               | Sheet Summary                              | Describe each sheet's purpose                                          |
|               | Logic Tracing                              | Show how values are derived                                            |
| Editing       | Prompt-Based Edits                         | "Add a column with profit %"                                           |
|               | Formula Generation                         | Generate and insert Excel formulas                                     |
|               | Multi-Step Commands                        | Handle chained actions like Rocode                                     |
|               | Structure Edits                            | Add/remove rows, tables, summaries                                     |
| UI            | Prompt Input                               | Sidebar chat with prompt entry                                         |
|               | AI Output Preview                          | Preview changes with \[Show Changes] and \[Apply] buttons              |
|               | History + Undo                             | Action history with undo/redo support                                  |
|               | Suggestions                                | Smart prompt suggestions refreshed dynamically                         |
| Accessibility | Keyboard Navigation, Screen Reader Support | Ensures compliance with WCAG and usability for users with disabilities |
| Onboarding    | Tutorial + Contextual Help                 | Built-in guide system with progressive disclosure                      |

---

### 3. User Stories (Sample)

* As a user, I can type a prompt to add a new column and preview changes before applying.
* As a user, I can ask "What does this formula do?" and get a human-readable explanation.
* As a user, I can undo an AI-made change via the action history panel.
* As a user, I can chain commands in a single prompt and apply changes one-by-one.
* As a user, I can save frequently used prompts for reuse.
* As a user, I can onboard quickly with tutorials and help prompts.
* As a user, I can navigate the tool using only a keyboard.

---

### 4. Success Metrics

| Metric                      | Target          |
| --------------------------- | --------------- |
| Prompt satisfaction rating  | >80%            |
| Prompt response time        | <5s             |
| Undo usage rate             | <10% of actions |
| Add-in install to first use | <1 min          |
| Action success rate         | >95%            |
| Accessibility compliance    | WCAG 2.1 AA     |
| AI prompt reuse rate        | >30%            |

---

## 📐 Technical Design Document (TDD)

### 1. System Architecture

```
        ┌────────────┐
        │ User Input │
        └─────┬──────┘
              ↓
      ┌────────────────┐
      │ Excel Add-in   │ ◄── HTML/JS UI + Office.js
      └─────┬────┬─────┘
            ↓    ↓
    ┌──────────────┐     ┌────────────────────┐
    │ Formula Parser│     │ Execution Engine   │
    └─────┬────────┘     └─────┬──────────────┘
          ↓                        ↓
      [AI Prompt Builder]     [Undo/History Stack]
              ↓
        ┌──────────────┐
        │ AI Gateway   │ ──► OpenAI / Claude / Local LLM
        └──────────────┘
            ↓
        ┌──────────────┐
        │ State Manager│ ◄── Workbook state, user prefs, AI context
        └──────────────┘
```

---

### 2. Component Breakdown (extended)

#### State Manager (extended)

```ts
interface DataManager {
  retentionPolicy: RetentionPolicy;
  backupStrategy: BackupStrategy;
  dataExport: DataExportService;
  versionControl: VersionControl;
}

interface RetentionPolicy {
  promptHistory: '30 days' | '90 days' | '1 year';
  undoStack: 'session' | '24 hours' | '7 days';
  userPreferences: 'persistent';
  analytics: 'anonymized 90 days';
}
```

#### Memory Manager

```ts
interface MemoryManager {
  chunkSize: number; // 1000 cells per chunk
  maxUndoStackSize: number; // 50 operations
  gcThreshold: number; // 80% memory usage
  largeFileStrategy: 'visible-only' | 'selected-range' | 'chunked';
}
```

#### UX Enhancements

```ts
interface ProgressiveEnhancement {
  offlineMode: OfflineCapabilities;
  performanceMode: PerformanceOptimizations;
  accessibilityMode: AccessibilityFeatures;
  enterpriseMode: EnterpriseFeatures;
}
```

#### Learning System

```ts
interface LearningSystem {
  userBehaviorAnalysis: BehaviorAnalyzer;
  promptOptimization: PromptOptimizer;
  personalizedSuggestions: SuggestionEngine;
  skillAssessment: SkillAssessor;
}
```

#### Collaboration Features

```ts
interface CollaborationFeatures {
  sharedPromptLibrary: SharedPrompts;
  teamTemplates: TeamTemplates;
  versionControl: CollaborativeVersioning;
  auditTrail: TeamAuditTrail;
}
```

---

### 11. Deployment Strategy

```yaml
Deployment:
  CI/CD Pipeline:
    - GitHub Actions or Azure DevOps
    - Automated testing on Excel Online
    - Staging environment validation
    - Blue-green deployment for zero downtime

  Environment Management:
    - Development: Local Excel + mock AI
    - Staging: Excel Online + test AI keys
    - Production: Excel Online + production AI
```

---

**Document Version:** v1.2
**Last Updated:** 2025-06-27 