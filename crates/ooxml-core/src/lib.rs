//! Infrastructure shared by Office Open XML formats (.docx, .xlsx, .pptx): the package (a ZIP
//! of parts) read with size limits and written back without losing anything, undo/redo over
//! its parts, relationships, `[Content_Types].xml` and small XML helpers. Nothing here knows
//! about WordprocessingML or SpreadsheetML.

pub mod content_types;
pub mod package;
pub mod rels;
pub mod xml;

pub use package::{Package, Undone, ZipLimits};
