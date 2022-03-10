# MS Word Formatter for Aviation Manuals

This VBA program was made while working to obtain an Air Operator Certificate for a startup airline.
The funding was unsuccessful.

After this fell through, I realised I liked coding and decided to pursue full stack engineering.

---

### **This program came from a task that quickly proved unachievable without automation:**

- Cleaning up and formatting drafts documents from multiple team members, who had wildly varying writing methods and limited knowledge of Word's best practices, and had modified imported existing manuals from existing airlines (breaking many things in the import);
- **Those are controlled manuals that easily reach thousands of pages**, and must adhere to strict regulations, have complex numbering systems, must have a complex header on each page, a table of contents, a list of effective pages with effective dates, etc.

---

## This program:

- Prompts the user to open a document to process, and an up to date .docx template for styles. The program document also has embedded templates (i.e. cover and preamble pages);
- [Prompts entry on a cover page mock-up](Screenshots/Cover.png) for later use - title, subtitle, date, version, authority, etc. - with data validation and sanitization;
- [Prompts user for an example of a table header](Screenshots/Headers.png), if present, for later recognition;
- Clears all bookmarks;
- Clears all section breaks, manual lne breaks, column breaks, and manual page breaks (except when a document orientation change is detected - portrait => landscape and vice versa);
-
