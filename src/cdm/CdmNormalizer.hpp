#pragma once
#include "document_model.hpp"

namespace cdm {

// Walk the entire document tree and merge consecutive Text inlines
// that share identical directStyle and styleRef. This dramatically
// reduces RTF output size for DOCX files that store per-character runs.
void Normalize(Document& doc);

} // namespace cdm
