# Outlook Smart Sort

Local-only Outlook VSTO add-in that predicts filing folders for email and supports single-message approval, grouped batch approval, undo, and retraining.

## Local Data

All add-in data stays on the local machine under:

`%LOCALAPPDATA%\OutlookClassifier`

The add-in stores:

- `store.sqlite`: training rows, explicit approval feedback, and decision history.
- `model.zip`: the ML.NET model.
- `model.meta.json`: model version, training counts, feature pipeline version, and evaluation summary.
- `logs\smartsort-yyyy-MM-dd.log`: local diagnostic logs.

No cloud services or external APIs are used.

## Privacy Controls

`FeedbackStore` exposes local-only maintenance methods for:

- clearing historical/scanned training rows and explicit feedback examples;
- clearing decisions and feedback history;
- configuring body snippet length through `BodySnippetLength` metadata;
- exporting diagnostics without subject/body content.

Diagnostic export includes counts and metadata only, not message subject or body text.

## Validation

Build and run the add-in from Visual Studio with Outlook desktop and VSTO installed. Non-Windows SDK builds may fail before C# compilation because `Microsoft.VisualStudio.Tools.Office.targets` is not available outside a VSTO install.
