// ImageInsert.h — shared deferred picture-insertion entry point (FxImage.cpp).
// Used by IMAGE and the barcode functions: queue a main-thread (WINDOW queue)
// insertion of an image file into the calling cell's rectangle.
#pragma once
#include <string>

namespace egtools::functions
{
    // fullAddr/cell: caller address parts (xloil::CallerInfo). source may be a
    // local file or http(s) URL (downloaded to a temp file). mode 0 keeps the
    // aspect ratio inside the cell. deleteLocalAfter: treat a LOCAL source file
    // as temporary and delete it once embedded (URLs are always temp-managed).
    void queueInsertPicture(const std::wstring& fullAddr, const std::wstring& cell,
                            const std::wstring& source, int mode, bool deleteLocalAfter);
}
