# AI Nonprofit Document Assistant

> **Empowering nonprofits with Generative AI to streamline document creation, unlock organizational knowledge, and increase mission impact.**

---

## Overview

The **AI Nonprofit Document Assistant** is an AI-powered Streamlit application designed to help nonprofit organizations reduce administrative burden, improve productivity, and accelerate document creation. By combining **Generative AI** with **Retrieval-Augmented Generation (RAG)**, the application enables nonprofit professionals to generate high-quality, mission-focused documents and interact with their organization's knowledge base through natural language.

The goal of this project is to make enterprise AI accessible to nonprofit organizations that often operate with limited staff, constrained budgets, and growing reporting requirements.

---

## The Problem

Nonprofit organizations spend countless hours creating grants, fundraising materials, board reports, strategic plans, policies, donor communications, and other mission-critical documents. At the same time, organizational knowledge is often scattered across shared drives, cloud storage, and multiple document formats, making it difficult to quickly locate accurate information.

These administrative challenges reduce productivity, slow decision-making, and take valuable time away from serving communities and advancing organizational missions.

---

## The Solution

The AI Nonprofit Document Assistant leverages Large Language Models (LLMs) and Retrieval-Augmented Generation (RAG) to automate document creation and provide intelligent access to organizational knowledge.

Users can:

- Generate professional nonprofit documents in minutes
- Chat with organizational documents
- Create consistent, mission-focused content
- Reduce repetitive administrative work
- Improve organizational knowledge sharing
- Support faster, data-informed decision-making

Instead of spending hours writing and searching for documents, nonprofit professionals can focus on creating greater impact in their communities.

---

# Features

## AI Document Generation

Generate more than **80 nonprofit-specific document types**, including:

### Fundraising & Development

- Grant Proposals
- Letters of Intent (LOIs)
- Case for Support
- Donor Thank-You Letters
- Annual Appeal Letters
- Sponsorship Proposals
- Capital Campaign Materials

### Executive Leadership

- Strategic Plans
- Annual Reports
- Executive Summaries
- Organizational Assessments
- Impact Reports

### Governance

- Board Agendas
- Board Meeting Minutes
- Board Resolutions
- Governance Policies
- Committee Charters

### Human Resources

- Employee Policies
- Volunteer Handbooks
- Job Descriptions
- Performance Reviews

### Communications

- Press Releases
- Newsletter Articles
- Website Content
- Social Media Posts
- Event Announcements

### Advocacy & Public Policy

- Policy Briefs
- Legislative Testimony
- White Papers
- Position Statements

### Research & Evaluation

- Needs Assessments
- Program Evaluations
- Logic Models
- Research Summaries

---

## AI Knowledge Assistant

The application includes Retrieval-Augmented Generation (RAG) capabilities that allow users to:

- Search organizational documents
- Ask natural language questions
- Retrieve relevant information instantly
- Receive AI-generated responses grounded in uploaded documents

### Example Questions

- What is our travel reimbursement policy?
- Summarize our strategic plan.
- What are the requirements in our latest grant agreement?
- What decisions were made during the last board meeting?

---

# Technology Stack

| Category | Technologies |
|----------|--------------|
| Frontend | Streamlit |
| Language | Python |
| AI Models | OpenAI GPT Models |
| Framework | LangChain |
| RAG | Vector Database, Semantic Search, Embeddings |
| Document Processing | PDF, DOCX, TXT |
| Version Control | Git & GitHub |

---

# Application Workflow

```text
User
 │
 ▼
Launch Streamlit Application
 │
 ▼
Enter OpenAI API Key
 │
 ▼
Upload Organizational Documents (Optional)
 │
 ▼
Select Document Template
 │
 ▼
Enter Prompt
 │
 ▼
OpenAI GPT + RAG Processing
 │
 ▼
Generate Professional Nonprofit Document
 │
 ▼
Review • Edit • Download
```

---

# Benefits

- Save hours of administrative work
- Accelerate grant writing
- Improve document consistency
- Preserve institutional knowledge
- Enhance organizational collaboration
- Support better decision-making
- Increase operational efficiency
- Allow staff to focus on mission delivery

---

# Current MVP

The current Streamlit MVP demonstrates the platform's core capabilities, including:

- AI-powered document generation
- Interactive chatbot interface
- Retrieval-Augmented Generation (RAG)
- Multiple nonprofit document templates
- User-friendly web interface

The MVP validates the concept while serving as the foundation for future enterprise deployment.

---

# Future Roadmap

## 🚀 Future Roadmap: AWS-Powered Commercial SaaS Platform

Mission Document Intelligence AI, LLC (MDI AI) is planning the evolution of its existing Streamlit-based AI Nonprofit Document Assistant into a secure, scalable, multi-tenant Software-as-a-Service (SaaS) platform hosted on **Amazon Web Services (AWS)**.

MDI AI is collaborating with **an AWS Premier Tier Services Partner**, to evaluate a phased cloud development strategy and pursue potential AWS funding support, subject to approval.

The proposed development roadmap consists of three phases:

### Phase 1: Technical Assessment & AWS Proof of Concept (POC)

Evaluate the existing Python/Streamlit application and establish the technical foundation for an AWS-native architecture.

**Planned capabilities:**
- Assess existing document ingestion, retrieval-augmented generation (RAG), summarization, and template-generation workflows.
- Evaluate **Amazon Bedrock** for managed foundation model inference.
- Explore **Amazon Bedrock Knowledge Bases** for document indexing, semantic retrieval, and source-grounded responses.
- Evaluate **Amazon Bedrock Guardrails** for AI safety and responsible model interactions.
- Develop an initial AWS proof of concept demonstrating core document intelligence functionality.
- Assess infrastructure deployment using **Terraform**.

### Phase 2: Secure Functional SaaS Pilot

Develop a secure, cloud-hosted application suitable for controlled nonprofit pilot deployments.

**Planned capabilities:**
- Implement authenticated nonprofit organizational workspaces.
- Establish tenant isolation to protect each nonprofit's documents and AI interactions.
- Introduce persistent document storage, indexing, and retrieval using appropriate AWS services.
- Support document uploads, AI-powered Q&A, summarization, source citations, and selected nonprofit document templates.
- Evaluate **Amazon Bedrock AgentCore** for orchestration of advanced AI workflows where appropriate.
- Implement foundational security controls, logging, monitoring, and document deletion workflows.
- Conduct functional, security, and user acceptance testing.

### Phase 3: Commercial SaaS Expansion

Expand the validated pilot application into a subscription-based AI document intelligence platform for nonprofit organizations.

**Planned capabilities:**
- Introduce subscription billing, service tiers, and usage enforcement.
- Expand and validate the library of 80+ nonprofit document templates.
- Develop integrations with external document repositories, including Microsoft SharePoint, OneDrive, and Google Drive.
- Enhance organization administration, user permissions, and document management.
- Expand AI-assisted document workflows and explore agentic capabilities through Amazon Bedrock services.
- Improve operational monitoring, scalability, deployment automation, and production readiness.
- Prepare the platform for broader commercial adoption by nonprofit organizations.

### ☁️ Proposed AWS Technology Stack

| AWS Technology | Proposed Function |
|---|---|
| Amazon Bedrock | Managed generative AI model inference |
| Amazon Bedrock Knowledge Bases | RAG-based document retrieval and grounding |
| Amazon Bedrock Guardrails | AI safety and content safeguards |
| Amazon Bedrock AgentCore | Potential orchestration of advanced AI workflows |
| Amazon S3 | Durable document and object storage |
| AWS IAM | Cloud resource access controls |
| Amazon CloudWatch | Application logging and operational monitoring |
| AWS Secrets Manager | Secure management of application secrets |
| Terraform | Infrastructure as Code and repeatable cloud deployments |

Additional AWS services for authentication, compute, networking, search, and billing integration will be selected during technical architecture assessment.

### 🎯 Long-Term Vision

MDI AI aims to make enterprise-grade AI document intelligence accessible to small and mid-sized nonprofit organizations without requiring extensive technical infrastructure or dedicated AI engineering teams.

The long-term goal is to provide a secure, subscription-based platform that helps nonprofits improve operational efficiency, streamline document preparation, retrieve institutional knowledge, and devote more resources to their missions.

**Development Status:** The existing Streamlit application serves as the functional MVP. The AWS architecture, implementation scope, development phases, and potential funding remain under evaluation with the AWS Premier Partner Services Partner and are subject to final technical assessment, contractual agreement, and AWS funding approval.
---

# Target Users

- Executive Directors
- Development Officers
- Grant Writers
- Fundraising Teams
- Communications Teams
- Program Managers
- Board Members
- Community Foundations
- Educational Nonprofits
- Faith-Based Organizations
- Policy & Advocacy Organizations

---

# Project Vision

My vision is to become the trusted AI operations platform for nonprofit organizations by helping them work smarter, make informed decisions faster, and maximize their mission impact through responsible, secure, and accessible artificial intelligence.

---

# Author

**Maxine Humphrey**

**DevOps Engineer | Enterprise AI | Cybersecurity Professional**

MBA • M.S. Business Analytics & Applied AI Candidate at the University of Rochester-Simon Business School

Founder, **Mission Document Intelligence AI**

---

# License

This project is intended for educational, research, and demonstration purposes. Please refer to the repository's LICENSE file for details.

---

# Acknowledgments

This project was made possible through the incredible open-source ecosystem, including:

- Streamlit
- OpenAI
- LangChain
- Python
- GitHub

Special thanks to the nonprofit professionals whose challenges inspired the development of this application.
