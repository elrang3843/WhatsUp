#include "FormatManager.h"
#include "TxtFormat.h"
#include "HtmlFormat.h"
#include "MdFormat.h"
#include "DocxFormat.h"
#include "HwpxFormat.h"
#include "HwpFormat.h"
#include "DocFormat.h"
#include "../i18n/Localization.h"
#include <memory>

FormatManager& FormatManager::Instance() {
    static FormatManager inst;
    return inst;
}

std::unique_ptr<IDocumentFormat> FormatManager::HandlerFor(DocFormat fmt) {
    switch (fmt) {
        case DocFormat::Txt:  return std::make_unique<TxtFormat>();
        case DocFormat::Html: return std::make_unique<HtmlFormat>();
        case DocFormat::Md:   return std::make_unique<MdFormat>();
        case DocFormat::Docx: return std::make_unique<DocxFormat>();
        case DocFormat::Hwpx: return std::make_unique<HwpxFormat>();
        case DocFormat::Hwp:  return std::make_unique<HwpFormat>();
        case DocFormat::Doc:  return std::make_unique<DocFormat_>();
        default:              return std::make_unique<TxtFormat>();
    }
}

FormatResult FormatManager::Load(const std::wstring& path, Document& doc) {
    DocFormat fmt = Document::DetectFormat(path);
    doc.SetFormat(fmt);

    auto handler = HandlerFor(fmt);
    if (!handler) {
        FormatResult r;
        r.error = Localization::Get(SID::MSG_FORMAT_UNSUPPORTED);
        return r;
    }
    return handler->Load(path, doc);
}

FormatResult FormatManager::Save(const std::wstring& path,
                                  const std::wstring& content,
                                  const std::string&  rtfContent,
                                  Document&           doc) {
    DocFormat fmt = Document::DetectFormat(path);
    doc.SetFormat(fmt);

    auto handler = HandlerFor(fmt);
    if (!handler) {
        FormatResult r;
        r.error = Localization::Get(SID::MSG_FORMAT_UNSUPPORTED);
        return r;
    }
    if (!handler->CanWrite()) {
        FormatResult r;
        r.error = Localization::Get(
            (fmt == DocFormat::Hwp) ? SID::MSG_HWP_READONLY : SID::MSG_DOC_READONLY);
        return r;
    }
    return handler->Save(path, content, rtfContent, doc);
}
