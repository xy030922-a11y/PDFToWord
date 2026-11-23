#pragma once

#include <QString>
#include <functional>

enum class ConversionMode { PreserveLayout, Editable };

using PageProgress = std::function<void(int completed, int total)>;

// An empty result means the converted file was saved successfully.
QString convertPdf(const QString &source, const QString &destination,
                   ConversionMode mode, const PageProgress &progress = {});
