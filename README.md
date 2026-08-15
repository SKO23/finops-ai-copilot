---
title: FinOps AI Copilot
emoji: 💼
colorFrom: indigo
colorTo: blue
sdk: docker
app_port: 7860
pinned: false
fullWidth: true
header: mini
short_description: AI expense review with Human-in-the-Loop
---

# 💼 FinOps AI Copilot

**AI-powered expense review with policy compliance, exception detection, evidence checking and Human-in-the-Loop decision making.**

FinOps AI Copilot helps Finance teams review employee expense claims against travel and expense policies.

The AI analyses uploaded documents and provides recommendations, but **the final Approve, Reject or Escalate decision remains with a human Finance reviewer.**

---

## 🎯 What it does

The application can:

- 📤 Upload expense claims, receipts, invoices and finance/travel policies
- 🧾 Extract and summarise expense information
- 📑 Check expenses against policy limits and rules
- ⚠️ Identify policy exceptions and potential risks
- 📎 Detect missing receipts, approvals or supporting evidence
- 🔎 Provide source document and page references
- 🤖 Generate an advisory AI recommendation
- 🤝 Record a Human-in-the-Loop Finance decision
- 📝 Capture reviewer notes
- 📄 Generate a downloadable Finance Review Report

---

## ⚡ Finance Quick Reviews

### 🧾 Expense Summary
Extracts key expense information such as:

- Expense type
- Date
- Amount
- Currency
- Merchant/provider
- Claimant
- Source document
- Page reference

### 📑 Policy Compliance
Compares claimed expenses with uploaded finance or travel policies and identifies:

- Compliant expenses
- Policy exceptions
- Expenses requiring manual review
- Relevant policy rules and limits

### ⚠️ Exceptions & Risks
Highlights material issues such as:

- Expenses above policy limits
- Non-reimbursable expenses
- Missing approvals
- Policy violations
- Items requiring further Finance review

### 📎 Missing Evidence
Identifies missing supporting documentation such as:

- Receipts
- Invoices
- Manager approvals
- Proof of payment
- Other evidence required by policy

---

## 🤖 AI Recommendation + Human-in-the-Loop

FinOps AI Copilot generates an **advisory AI recommendation** based on the uploaded documents.

Example recommendations include:

- APPROVE
- REJECT
- ESCALATE / MANUAL REVIEW

The AI does **not** make the final Finance decision.

An authorised reviewer records the final outcome using:

- ✅ Approve
- ❌ Reject
- 👤 Escalate / Manual Review

The reviewer can also enter a note explaining the final decision.

This keeps the application aligned with a **Human-in-the-Loop governance model**.

---

## 👁️ Document and receipt extraction

The application uses a two-stage extraction approach.

### Native PDF text

For text-based PDFs, the application first uses standard PDF text extraction.

This is fast and preserves document/page information.

### Gemini vision / OCR fallback

If readable text cannot be extracted from a document, the application can use Gemini's multimodal document understanding as a fallback.

This enables the MVP to process:

- Scanned PDFs
- Photographed receipts
- JPG/JPEG files
- PNG files
- WEBP images

The extracted content is then passed into the same Finance review workflow.

---

## 🔎 Grounded document analysis

The application uses document retrieval to help ground answers in uploaded evidence.

The workflow combines:

- PDF/document extraction
- Text chunking
- Sentence Transformer embeddings
- Semantic retrieval
- Gemini-based analysis
- File and page references

Finance reviewers can also ask questions such as:

> Is the hotel expense within policy?

> Which expenses require manual review?

> What supporting evidence is missing?

> Why is the taxi claim being escalated?

---

## 🏗️ High-level architecture

```text
                    ┌─────────────────────┐
                    │ Expense / Policy    │
                    │ Documents           │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │ Document Extraction │
                    │                     │
                    │ PDF text first      │
                    │ Vision/OCR fallback │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │ Chunking +          │
                    │ Embeddings          │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │ Evidence Retrieval  │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │ Gemini Finance      │
                    │ Analysis            │
                    └──────────┬──────────┘
                               │
          ┌────────────────────┼────────────────────┐
          │                    │                    │
          ▼                    ▼                    ▼
   Policy Compliance     Exceptions & Risks    Missing Evidence
          │                    │                    │
          └────────────────────┼────────────────────┘
                               │
                    ┌──────────▼──────────┐
                    │ AI Recommendation   │
                    │ Advisory Only       │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │ Human Finance       │
                    │ Decision            │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │ Finance Review      │
                    │ Report              │
                    └─────────────────────┘
```

---

## 🛠️ Technology stack

- **Python**
- **Gradio**
- **Google Gemini API**
- **Sentence Transformers**
- **all-MiniLM-L6-v2 embeddings**
- **PyPDF**
- **python-docx**
- **Docker**
- **Hugging Face Spaces**

---

## 📄 Finance Review Report

The application can generate an editable Word report containing:

- Finance case summary
- Documents reviewed
- Expense Summary
- Policy Compliance
- Exceptions & Risks
- Missing Evidence
- Source/page evidence
- AI recommendation
- Human Finance decision
- Reviewer notes
- AI governance notice

This provides an auditable separation between the **AI recommendation** and the **human decision**.

---

## 🧪 Synthetic test scenario

The MVP has been tested using synthetic travel-expense documents containing:

- Hotel expense above the permitted nightly limit
- Compliant meal expense
- Taxi expense with a missing receipt
- Non-reimbursable alcohol expense

The Copilot identifies the relevant policy exceptions and supporting evidence before generating an advisory recommendation for Finance review.

No real employee or confidential financial information is required for the demo.

---

## 🔐 Security

The Gemini API key is stored as a **Hugging Face Space Secret**.

API keys and credentials should never be committed to the repository.

---

## ⚠️ Important notice

This application is an **MVP / demonstration project**.

AI-generated analysis may contain errors and should not be treated as an automatic financial approval or rejection.

Material findings should be checked against the cited source documents.

**The AI recommendation is advisory only. The authorised Finance reviewer retains the final decision.**

---

## 🚀 Future roadmap

Potential future enhancements include:

- ERP and expense-management system integration
- SAP / Oracle / Workday integrations
- Automated expense ingestion
- Corporate card reconciliation
- More advanced receipt extraction
- Fraud and duplicate-claim detection
- Organisation-specific policy libraries
- Role-based access controls
- Audit logs
- Persistent case storage
- Finance workflow integration
- Approval routing
- Enterprise monitoring and governance

---

## 👤 Project

Built as an AI / FinOps MVP demonstrating:

**Document AI + Retrieval + Finance Policy Reasoning + Multimodal OCR + Human-in-the-Loop Governance**
