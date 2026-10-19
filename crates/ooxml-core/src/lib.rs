//! Infrastructure shared by Office Open XML formats (.docx, .xlsx, .pptx): the package (a ZIP
//! of parts) read with size limits and written back without losing anything, undo/redo over
//! its parts, relationships, `[Content_Types].xml` and small XML helpers. Nothing here knows
//! about WordprocessingML or SpreadsheetML. `markers` holds the shared lexicon of template
//! fields (`{{KEY}}`, `{key}`) and pending content (`[ … ]`).

pub mod content_types;
pub mod markers;
pub mod package;
pub mod rels;
pub mod xml;

pub use package::{Package, Undone, ZipLimits};
