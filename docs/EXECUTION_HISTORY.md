# Execution History & Metrics

This document provides a high-level summary of the historical performance and large-scale operations conducted using this toolkit.

## 📈 Historical Execution Summary

### Large-scale mailbox operations (Jan 2026)
*   **Operations:** Bulk mailbox exports, folder structure standardization, and message migration.
*   **Scale:** 
    *   Bulk mailbox exports, folder-structure standardization, and message migration across many mailboxes via the Microsoft Graph API.
*   **Performance Metrics:**
    *   **Resilience:** The `Invoke-GraphWithRetry` helper successfully navigated HTTP 429 (Throttling) and 5xx (Server Error) responses during peak load.
    *   **Data Integrity:** Automated mailbox folder verification confirmed zero data loss during move operations between project folder hierarchies.

## 🛠️ Lessons Learned
*   **Graph Throttling:** Large-scale migrations work best when sessions are distributed or utilize exponential backoff (now standard in `GraphHelpers.ps1`).
*   **Mailbox Depth:** Recursively searching folders can be expensive; the `Get-MailFoldersWithPaths` function with progress reporting was essential for maintaining visibility.
*   **Protected Folders:** System folders (Inbox, Sent Items, etc.) require specific protection logic to prevent accidental modification during "Ensure" modes.

---
*Note: this summary is deliberately generic. No tenant data, mailbox lists, logs, or individual addresses are included in this repository; examples throughout use fictional addresses.*
