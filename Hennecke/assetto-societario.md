# Hennecke GROUP — Nuovo assetto societario

> **Aggiornamento:** acquisizione del Gruppo Hennecke da parte del **Brückner Group SE**.
> Accordo firmato il **3 dicembre 2025**, perfezionamento del closing il **5 gennaio 2026**.
> Il venditore era un fondo gestito da **Capvis AG** (Svizzera), in precedenza azionista di maggioranza.

Con questa operazione il Gruppo Hennecke entra a far parte del Brückner Group SE, di cui
diventa la quinta divisione (business unit) accanto a Brückner Maschinenbau, Brückner Servtec,
Kiefel e PackSys Global. La struttura interna del Gruppo Hennecke (i due brand di prodotto e
le società controllate) rimane invariata; cambia la proprietà al vertice.

---

## 1. Diagramma del nuovo assetto societario

```mermaid
flowchart TD
    BG["<b>Brückner Group SE</b><br/>Siegsdorf, Germania<br/><i>Capogruppo / Holding</i>"]

    BG --> BM["Brückner<br/>Maschinenbau"]
    BG --> BS["Brückner<br/>Servtec"]
    BG --> KF["Kiefel"]
    BG --> PG["PackSys Global"]
    BG --> HG["<b>Hennecke GROUP</b><br/>HQ Sankt Augustin, DE<br/><i>~680 dipendenti · ~15 società</i><br/>(nuova divisione, dal 2026)"]

    %% --- Struttura interna del Gruppo Hennecke ---
    HG --> B1["<b>Brand: Hennecke<br/>Polyurethane Technology</b>"]
    HG --> B2["<b>Brand: Hennecke-OMS</b>"]

    B1 --> H_DE["Hennecke GmbH<br/>Sankt Augustin, DE"]
    B1 --> H_US["Hennecke Inc.<br/>Pittsburgh, USA"]
    B1 --> H_CN1["Hennecke Machinery<br/>(Shanghai) Ltd. — CN"]
    B1 --> H_CN2["Hennecke Machinery<br/>(Jiaxing) Ltd. — CN"]
    B1 --> H_AS["Hennecke Asia Pte. Ltd.<br/>Singapore<br/><i>(uffici: TH · IN · VN)</i>"]
    B1 --> H_JP["Hennecke Machinery<br/>Japan K.K. — Tokyo"]
    B1 --> H_KR["Hennecke Korea Ltd.<br/>Seoul"]

    B2 --> O_IT["Hennecke-OMS S.p.A.<br/>Verano Brianza, IT"]

    classDef holding fill:#1f3a5f,stroke:#0d1b2a,color:#ffffff,stroke-width:2px;
    classDef hennecke fill:#c8102e,stroke:#7a0a1c,color:#ffffff,stroke-width:2px;
    classDef brand fill:#f2c200,stroke:#b38f00,color:#1a1a1a,stroke-width:1px;
    classDef sub fill:#f5f5f5,stroke:#999999,color:#1a1a1a;
    classDef sibling fill:#e6edf5,stroke:#6b86a8,color:#1a1a1a;

    class BG holding;
    class HG hennecke;
    class B1,B2 brand;
    class BM,BS,KF,PG sibling;
    class H_DE,H_US,H_CN1,H_CN2,H_AS,H_JP,H_KR,O_IT sub;
```

---

## 2. Cronologia degli assetti proprietari (catena delle cessioni)

```mermaid
flowchart LR
    A["<b>Bayer<br/>MaterialScience</b><br/>(origine)"]
    -->|"2008<br/>cessione"| B["<b>ADCURAM<br/>Group AG</b><br/>Monaco di Baviera"]
    -->|"apr. 2016<br/>cessione"| C["<b>Capvis</b><br/>(Capvis Equity IV)<br/>Svizzera"]
    -->|"closing 5 gen. 2026<br/><b>ultima cessione</b>"| D["<b>Brückner<br/>Group SE</b><br/>Siegsdorf, DE"]

    classDef past fill:#f5f5f5,stroke:#999,color:#1a1a1a;
    classDef current fill:#1f3a5f,stroke:#0d1b2a,color:#fff,stroke-width:2px;
    class A,B,C past;
    class D current;
```

---

## 3. Note di sintesi

| Voce | Dettaglio |
|---|---|
| **Nuovo proprietario** | Brückner Group SE (Siegsdorf, Germania) |
| **Quota acquisita** | 100% del Gruppo Hennecke |
| **Venditore** | Fondo gestito da Capvis AG (Svizzera) |
| **Firma accordo** | 3 dicembre 2025 |
| **Closing** | 5 gennaio 2026 |
| **Posizionamento** | Hennecke = nuova/quinta divisione del Brückner Group |
| **Sede Gruppo Hennecke** | Sankt Augustin, Germania |
| **Brand di prodotto** | Hennecke Polyurethane Technology · Hennecke-OMS |
| **Dimensione** | ~680 dipendenti, ~15 società nel mondo |
| **Settore** | Macchine e impianti per la lavorazione del poliuretano (PUR) |

### Fonti
- [Brückner Group acquires Hennecke Group — PU MAGAZINE](https://www.pu-magazine.com/pu/news/meldungen/20251212-brueckner-acquires-hennecke.php)
- [CMS advises Brückner Group on the acquisition of Hennecke Group](https://cms.law/en/deu/news-information/cms-advises-brueckner-group-on-the-acquisition-of-hennecke-group)
- [Hengeler Mueller advises Capvis on sale of Hennecke Group to Brückner Group](https://hengeler-news.com/en/articles/hengeler-mueller-advises-capvis-on-sale-of-hennecke-group-to-brueckner-group)
- [Hennecke-OMS S.p.A. — Hennecke GROUP](https://www.hennecke.com/en/company/worldwide/hennecke-oms)
- [About Us — Brückner Group](https://www.brueckner.com/en/about-us)
- [Adcuram: Verkauf der Hennecke-Gruppe an Finanzinvestor Capvis (2016)](https://www.k-online.de/de/Media_News/News/Archiv_Branchen-News/Adcuram_Verkauf_der_Hennecke-Gruppe_an_Finanzinvestor_Capvis)
