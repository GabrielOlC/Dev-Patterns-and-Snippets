# Dev-Patterns-and-Snippets

**A centralized library establishing my naming conventions, architectural principles and reusable code snippets (VBA, M, Excel, SQL, JavaScript, Python). It ensures that every project I build is modular, maintainable, and scalable from Day 1.**

---

## 🏗️ Philosophy

> "Reliability and maintainability are not accidents; they are the result of rigorous standards!"

This repository documents my personal framework for consistency across all my projects. It separates Interface, Logic and Data into clear layers (`wf_` controllers, `cm_` services, `sys_` core), applies **Object Calisthenics** and **Self-Documenting Code**, and builds every reusable piece once, so the next project starts from proven code instead of an empty module.

It is important to note that this is a Living Doc: standards and naming conventions may change over time, as I keep learning from each project.

## 🧰 The Toolkit

Services that prove useful beyond one project live in `/.Snippets`, each opening with the header from `/Templates`. `service-catalog.md` lists every one of them in a line, so it's quick to check what already exists before writing something new.

## 📂 Repository Structure

```markup-templating
/Dev-Patterns-and-Snippets
│
├── service-catalog.md             # One line per service in the toolkit
│
├── /Standards
│   ├── Naming_Conventions.md      # Tokens and casing (sys_, bv_, fn, tb_, wf_/cm_ files...)
│   └── Architecture_Principles.md # Layers, reuse, data integrity, error handling, documentation
│
├── /.Snippets                     # The toolkit: reusable services
│   ├── /VBA                       # Reusable modules (Table select/resizing, Regex to RGB...)
│   ├── /PowerQuery_M              # Data cleaning functions (Unzip, Parse XML...)
│   ├── /Excel                     # Advanced formula logic (lmdNotEmptyOrZero...)
│   └── /Python                    # Utility scripts (Scrapers, Parsers...)
│
└── /Templates
    └── In code Snippets.md        # The header every service file opens with
```
