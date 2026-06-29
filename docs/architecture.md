---
title: Architecture
---

# Architecture

This page shows how the modules collaborate, the lifecycle of a single run, and
the design principles that shaped the project.

## Design principles

1. **One job per module.** Each file has a single responsibility, so it can be
   read, tested and changed in isolation. See the [Modules overview](modules/index.md).
2. **Template lives in Outlook, not in code.** Your formatting, images and
   signature come from a real draft — code only edits three fields.
3. **Fail safe, continue on row errors.** Setup problems stop the run; a single
   bad row is logged and skipped.
4. **Drafts only, always.** No `.Send()` anywhere.
5. **Configuration over hard-coding.** Everything tunable lives in
   [`config.py`](config.md) and can be overridden by environment variables.

## Module dependency map

```mermaid
flowchart TD
    main[main.py<br/><i>orchestrator</i>]
    config[config.py]
    logger[logger.py]
    reader[excel_reader.py]
    models[models.py]
    outlook[outlook_service.py]
    html[html_table_generator.py]
    exc[exceptions.py]

    main --> config
    main --> logger
    main --> reader
    main --> outlook
    reader --> models
    reader --> exc
    reader --> logger
    outlook --> html
    outlook --> models
    outlook --> exc
    outlook --> logger

    classDef entry fill:#5e35b1,stroke:#311b92,color:#fff;
    classDef pure fill:#1e88e5,stroke:#0d47a1,color:#fff;
    class main entry;
    class html,models pure;
```

- **Purple** = entry point.
- **Blue** = pure, dependency-light modules (easy to unit-test without Outlook).

!!! note "Why `models` and `html_table_generator` have no project dependencies"
    They are pure data/string logic. Keeping them free of Outlook/Excel imports
    means you can test them on any OS and reuse them anywhere.

## Lifecycle of a run

```mermaid
sequenceDiagram
    autonumber
    participant U as User / Task Scheduler
    participant M as main.py
    participant C as config.py
    participant R as excel_reader.py
    participant O as outlook_service.py
    participant H as html_table_generator.py
    participant X as Outlook (COM)

    U->>M: python -m src.main
    M->>C: load_config()
    C-->>M: AppConfig
    M->>M: setup_logging(log_dir)
    M->>R: read(excel_path)
    R-->>M: list[Contact]
    M->>O: connect()
    O->>X: Dispatch("Outlook.Application")
    M->>O: locate_template()
    O->>X: find draft + SaveAs(.oft)
    loop each Contact
        M->>O: create_draft_for(contact)
        O->>X: CreateItemFromTemplate(.oft)
        O->>H: generate_html_table(values)
        H-->>O: <table>…</table>
        O->>X: set To/Cc/Subject/HTMLBody + Save()
        O-->>M: new EntryID
    end
    M-->>U: summary + exit code
```

## Data transformation pipeline

How one spreadsheet row becomes a draft:

```mermaid
flowchart LR
    subgraph Excel
      R[("Row:<br/>514 | 1526 | … | note")]
    end
    R --> CT[Contact.values]
    CT -->|values[0]| TO[To = 514]
    CT -->|values[:7] joined| SUB[Subject]
    CT -->|all values| TAB[generate_html_table]
    TAB --> INS{{TABLE placeholder?}}
    INS -->|yes| REP[replace in body]
    INS -->|no| APP[insert before &lt;/body&gt;]
    REP & APP --> BODY[HTMLBody]
    TO & SUB & BODY --> SAVE[(Save draft)]
```

See the exact rules in the [Data contract](reference/data-contract.md).

## Error-handling strategy

```mermaid
flowchart TD
    Start([run]) --> Cfg[load config + logging]
    Cfg --> Read{read Excel}
    Read -- WorkbookError --> Stop1[[log + exit 1]]
    Read -- ok --> Conn{connect + locate template}
    Conn -- Outlook/Template error --> Stop2[[log + exit 1]]
    Conn -- ok --> Loop[for each row]
    Loop --> Try{create_draft_for}
    Try -- success --> Count[created++]
    Try -- EmailAutomationError --> Skip[log + failed++]
    Try -- unexpected Exception --> Skip
    Count & Skip --> More{more rows?}
    More -- yes --> Loop
    More -- no --> Sum[[summary + exit 0/2]]
```

| Exit code | Meaning |
| --- | --- |
| `0` | All rows succeeded |
| `1` | Fatal setup error (bad workbook, Outlook, or template) |
| `2` | Completed, but one or more rows failed |

Read the deeper rationale in [`main.py`](main.md) and [`exceptions.py`](exceptions.md).
