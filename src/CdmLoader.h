#pragma once
#include <windows.h>
#include "cdm/document_model.hpp"

class Editor;

// Apply a cdm::Document to a RichEdit control directly via
// SetText + EM_SETCHARFORMAT + EM_SETPARAFORMAT.
// No RTF is generated; no EM_STREAMIN is used.
// textColor: theme text color (pass Application::Instance().TextColor())
void LoadCdmDocument(const cdm::Document& doc,
                     Editor*              editor,
                     COLORREF             textColor = RGB(0, 0, 0));
