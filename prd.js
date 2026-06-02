const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Footer, Header, TabStopType, TabStopPosition,
  VerticalAlign
} = require('docx');
const fs = require('fs');

const PURPLE_DARK = "3C3489";
const PURPLE_MID = "534AB7";
const PURPLE_LIGHT = "EEEDFE";
const PURPLE_BORDER = "AFA9EC";
const TEAL = "0F6E56";
const TEAL_LIGHT = "E1F5EE";
const GRAY = "5F5E5A";
const GRAY_LIGHT = "F1EFE8";
const WHITE = "FFFFFF";

const border = { style: BorderStyle.SINGLE, size: 1, color: "CECBF6" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: PURPLE_MID, space: 4 } },
    children: [new TextRun({ text, font: "Arial", size: 32, bold: true, color: PURPLE_DARK })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 260, after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: PURPLE_MID })]
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "2C2C2A", ...opts })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: "2C2C2A" })]
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun("")] });
}

function infoBox(label, value) {
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 2500, type: WidthType.DXA },
        shading: { fill: PURPLE_LIGHT, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: label, font: "Arial", size: 20, bold: true, color: PURPLE_DARK })] })]
      }),
      new TableCell({
        borders,
        width: { size: 6860, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 160, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: value, font: "Arial", size: 20, color: "2C2C2A" })] })]
      })
    ]
  });
}

function sectionTable(headers, rows, colWidths) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      borders,
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: PURPLE_MID, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 160, right: 120 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 20, bold: true, color: WHITE })] })]
    }))
  });

  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders,
      width: { size: colWidths[ci], type: WidthType.DXA },
      shading: { fill: ri % 2 === 0 ? WHITE : GRAY_LIGHT, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 160, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 20, color: "2C2C2A" })] })]
    }))
  }));

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows]
  });
}

function statusBadge(status) {
  const colors = { "Must have": PURPLE_MID, "Should have": "0F6E56", "Nice to have": "BA7517" };
  return colors[status] || GRAY;
}

const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: PURPLE_DARK },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: PURPLE_MID },
        paragraph: { spacing: { before: 260, after: 120 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: PURPLE_MID, space: 4 } },
            spacing: { after: 120 },
            children: [
              new TextRun({ text: "Time Log Entry Form", font: "Arial", size: 18, bold: true, color: PURPLE_MID }),
              new TextRun({ text: "   |   Product Requirements Document", font: "Arial", size: 18, color: GRAY })
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 6, color: PURPLE_BORDER, space: 4 } },
            spacing: { before: 120 },
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            children: [
              new TextRun({ text: "Confidential - Internal Use Only", font: "Arial", size: 18, color: GRAY }),
              new TextRun({ text: "\tPage ", font: "Arial", size: 18, color: GRAY }),
              new TextRun({ children: [new PageNumber()], font: "Arial", size: 18, color: GRAY })
            ]
          })
        ]
      })
    },
    children: [
      new Paragraph({
        spacing: { before: 0, after: 0 },
        shading: { fill: PURPLE_MID, type: ShadingType.CLEAR },
        children: [new TextRun({ text: "" })]
      }),
      spacer(),
      new Paragraph({
        spacing: { before: 80, after: 40 },
        children: [new TextRun({ text: "Product Requirements Document", font: "Arial", size: 48, bold: true, color: PURPLE_DARK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: "Time Log Entry Form", font: "Arial", size: 36, color: PURPLE_MID })]
      }),
      spacer(),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2500, 6860],
        rows: [
          infoBox("Document version", "1.0"),
          infoBox("Status", "Draft"),
          infoBox("Author", "Product Team"),
          infoBox("Date", new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })),
          infoBox("Audience", "Engineering, Design, QA"),
        ]
      }),

      spacer(), spacer(),
      h1("1. Overview"),
      h2("1.1 Purpose"),
      body("This document defines the product requirements for the Time Log Entry Form, a web-based interface that enables users to record time-tracked activities with supporting proof. It serves as the single source of truth for design, engineering, and QA teams throughout the development lifecycle."),
      spacer(),

      h2("1.2 Background"),
      body("Teams tracking billable hours, task durations, or compliance-related time events currently lack a standardised, verifiable submission mechanism. The Time Log Entry Form addresses this gap by combining identity capture, precise duration logging, locale awareness, and photographic proof into a single, streamlined interface."),
      spacer(),

      h2("1.3 Goals"),
      bullet("Provide a simple, accessible form for logging time entries with proof."),
      bullet("Ensure accurate duration capture with hours, minutes, and seconds granularity."),
      bullet("Support locale/timezone awareness for globally distributed teams."),
      bullet("Enable photo upload as verifiable proof of task completion."),
      bullet("Deliver a polished, responsive UI that reduces friction in submission."),
      spacer(),

      h1("2. User Personas"),
      body("The form targets two primary user types:"),
      spacer(),

      sectionTable(
        ["Persona", "Role", "Primary Need"],
        [
          ["Field Contributor", "Individual submitter (staff, contractor, freelancer)", "Quick, accurate time entry with proof upload from any device"],
          ["Team Manager", "Reviewer of submitted logs", "Confidence that entries are complete, verified, and timestamped"],
          ["Compliance Officer", "Auditor of records", "Exportable, structured data with attached proof for audit trails"],
        ],
        [2600, 3760, 3000]
      ),
      spacer(), spacer(),

      h1("3. Functional Requirements"),
      h2("3.1 Form fields"),
      body("The form must include the following fields:"),
      spacer(),

      sectionTable(
        ["Field", "Type", "Validation", "Priority"],
        [
          ["Full name", "Text input", "Required, non-empty, max 100 chars", "Must have"],
          ["Locale / timezone", "Dropdown select", "Required, valid IANA timezone string", "Must have"],
          ["Hours", "Number input", "Required*, min 0, max 99", "Must have"],
          ["Minutes", "Number input", "Required*, min 0, max 59, auto-clamp", "Must have"],
          ["Seconds", "Number input", "Required*, min 0, max 59, auto-clamp", "Must have"],
          ["Proof photo", "File upload", "Optional, image files only, max 10MB", "Should have"],
        ],
        [2000, 2000, 3160, 2200]
      ),
      spacer(),
      body("* At least one of hours, minutes, or seconds must be greater than zero."),
      spacer(),

      h2("3.2 Duration display"),
      bullet("A live summary (e.g. \"2h 30m 45s\") must update in real time as the user types into any time field."),
      bullet("Minutes and seconds inputs must auto-clamp to their valid ranges (0-59) on blur."),
      bullet("The total duration display must always reflect the latest valid values."),
      spacer(),

      h2("3.3 Photo upload"),
      bullet("Accepted formats: PNG, JPG, WEBP."),
      bullet("Maximum file size: 10MB. Files exceeding this must be rejected with a clear error message."),
      bullet("On selection, a preview thumbnail of the image must render within the form."),
      bullet("A remove button must allow the user to deselect the file and return to the upload prompt."),
      bullet("Drag-and-drop must be supported in addition to the click-to-browse interaction."),
      spacer(),

      h2("3.4 Validation and error handling"),
      bullet("All required fields must be validated on submit."),
      bullet("Inline, contextual error messages must appear near the relevant field."),
      bullet("The submit button must remain active but trigger validation on click."),
      bullet("Error messages must auto-dismiss after 3.5 seconds or on field interaction."),
      spacer(),

      h2("3.5 Submission and confirmation"),
      bullet("On successful submission, the form view must transition to a confirmation screen."),
      bullet("The confirmation screen must display a summary of the submitted values: name, locale, duration, filename (if uploaded), and submission timestamp."),
      bullet("A reset action must clear all fields and return the user to the empty form."),
      spacer(),

      h1("4. Non-Functional Requirements"),
      spacer(),
      sectionTable(
        ["Category", "Requirement"],
        [
          ["Performance", "Form must be interactive within 2 seconds on a standard 4G connection."],
          ["Accessibility", "All form fields must be keyboard-navigable with descriptive ARIA labels."],
          ["Responsiveness", "Layout must adapt gracefully for screen widths from 320px to 1440px."],
          ["Browser support", "Must function on the latest two versions of Chrome, Firefox, Safari, and Edge."],
          ["Security", "File uploads must be validated client-side for type and size before any server submission."],
          ["Internationalisation", "Locale dropdown must cover major IANA timezone regions globally."],
        ],
        [2400, 6960]
      ),
      spacer(), spacer(),

      h1("5. User Interface Specifications"),
      h2("5.1 Layout"),
      body("The form is presented as a single-page card layout centred on the viewport. It is divided into three logical sections separated by visual dividers:"),
      bullet("Personal info: full name and locale fields."),
      bullet("Duration: hours, minutes, and seconds inputs within a highlighted time block, with the live summary display."),
      bullet("Proof photo: upload zone with drag-and-drop support and image preview."),
      spacer(),

      h2("5.2 Visual design"),
      body("The interface uses a purple-dominant colour palette to signal trust and professionalism. Key design tokens:"),
      spacer(),
      sectionTable(
        ["Element", "Colour / Style"],
        [
          ["Card header", "Deep purple (#534AB7) background with light purple (#EEEDFE) text"],
          ["Section labels", "Purple accent (#7F77DD), uppercase, tracked lettering"],
          ["Input borders", "Light purple (#CECBF6), transitioning to deep purple on focus"],
          ["Focus ring", "3px solid light purple (#EEEDFE)"],
          ["Submit button", "Deep purple (#534AB7) fill, light purple text, darkens on hover"],
          ["Time block", "Light purple (#EEEDFE) background with purple border"],
          ["Success screen", "Teal (#1D9E75) border and iconography on white card"],
          ["Error messages", "Red (#FCEBEB) background with red border (#F09595)"],
        ],
        [3000, 6360]
      ),
      spacer(),

      h2("5.3 Interaction states"),
      bullet("Empty state: placeholder text visible in all inputs; upload zone shows icon and hint text."),
      bullet("Active/focused state: input border transitions to deep purple with a focus ring."),
      bullet("Filled state: user-entered values displayed clearly; upload zone shows image preview."),
      bullet("Error state: red banner below the form with a descriptive message."),
      bullet("Success state: form replaced by a teal-accented confirmation card with a full submission summary."),
      spacer(),

      h1("6. Out of Scope"),
      body("The following are explicitly excluded from this version of the product:"),
      bullet("Backend API integration or data persistence."),
      bullet("Authentication or user account management."),
      bullet("Multi-step or wizard-style form flow."),
      bullet("Bulk submission or CSV import."),
      bullet("Admin dashboard or log review interface."),
      bullet("Native mobile application wrapper."),
      spacer(),

      h1("7. Success Metrics"),
      body("The following metrics will be used to evaluate the success of this feature post-launch:"),
      spacer(),
      sectionTable(
        ["Metric", "Target", "Measurement method"],
        [
          ["Form completion rate", ">= 85%", "Analytics: submit events / form load events"],
          ["Validation error rate", "< 15% of submissions", "Analytics: error trigger events"],
          ["Time to complete", "< 90 seconds median", "Session recording analysis"],
          ["Photo upload adoption", ">= 60% of submissions include a photo", "Backend upload logs"],
          ["User satisfaction (CSAT)", ">= 4.0 / 5.0", "Post-submission feedback prompt"],
        ],
        [2800, 2400, 4160]
      ),
      spacer(), spacer(),

      h1("8. Open Questions"),
      body("The following questions require resolution before engineering begins:"),
      spacer(),
      sectionTable(
        ["#", "Question", "Owner", "Status"],
        [
          ["1", "Where are submitted entries persisted? (Database schema, API contract)", "Engineering Lead", "Open"],
          ["2", "Are photo uploads stored in object storage (e.g. S3) or inline as base64?", "Engineering Lead", "Open"],
          ["3", "Should locale auto-detect from browser, with manual override?", "Product Manager", "Open"],
          ["4", "Is a submission confirmation email required?", "Product Manager", "Open"],
          ["5", "What is the data retention policy for uploaded proof photos?", "Legal / Compliance", "Open"],
        ],
        [600, 4200, 2400, 2160]
      ),
      spacer(), spacer(),

      h1("9. Revision History"),
      spacer(),
      sectionTable(
        ["Version", "Date", "Author", "Notes"],
        [
          ["1.0", new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }), "Product Team", "Initial draft"],
        ],
        [1200, 2400, 2760, 3000]
      ),
      spacer(),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/time-log-form-prd.docx', buffer);
  console.log('Done');
});
