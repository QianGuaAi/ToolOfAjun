#include "ui/renderer_d2d.h"

#include <algorithm>

namespace mytools {
namespace {

D2D1_COLOR_F Rgb(float r, float g, float b, float a = 1.0f) {
    return D2D1::ColorF(r / 255.0f, g / 255.0f, b / 255.0f, a);
}

float Scale(float value, float dpi_scale) {
    return value * dpi_scale;
}

}  // namespace

bool RendererD2D::Initialize() {
    HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, d2d_factory_.GetAddressOf());
    if (FAILED(hr)) {
        return false;
    }

    hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED,
                             __uuidof(IDWriteFactory),
                             reinterpret_cast<IUnknown**>(dwrite_factory_.GetAddressOf()));
    if (FAILED(hr)) {
        return false;
    }

    dwrite_factory_->CreateTextFormat(L"Segoe UI",
                                      nullptr,
                                      DWRITE_FONT_WEIGHT_SEMI_BOLD,
                                      DWRITE_FONT_STYLE_NORMAL,
                                      DWRITE_FONT_STRETCH_NORMAL,
                                      24.0f,
                                      L"zh-cn",
                                      title_format_.GetAddressOf());

    dwrite_factory_->CreateTextFormat(L"Segoe UI",
                                      nullptr,
                                      DWRITE_FONT_WEIGHT_NORMAL,
                                      DWRITE_FONT_STYLE_NORMAL,
                                      DWRITE_FONT_STRETCH_NORMAL,
                                      14.0f,
                                      L"zh-cn",
                                      body_format_.GetAddressOf());

    dwrite_factory_->CreateTextFormat(L"Segoe UI",
                                      nullptr,
                                      DWRITE_FONT_WEIGHT_NORMAL,
                                      DWRITE_FONT_STYLE_NORMAL,
                                      DWRITE_FONT_STRETCH_NORMAL,
                                      12.0f,
                                      L"zh-cn",
                                      status_format_.GetAddressOf());

    if (title_format_) {
        title_format_->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
    }
    if (body_format_) {
        body_format_->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
    }
    if (status_format_) {
        status_format_->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
    }

    return title_format_ && body_format_ && status_format_;
}

void RendererD2D::DiscardDeviceResources() {
    render_target_.Reset();
    text_brush_.Reset();
    muted_brush_.Reset();
    accent_brush_.Reset();
    border_brush_.Reset();
}

void RendererD2D::Resize(HWND window) {
    if (!render_target_) {
        return;
    }

    RECT rect{};
    GetClientRect(window, &rect);
    const D2D1_SIZE_U size =
        D2D1::SizeU(static_cast<UINT32>(rect.right - rect.left), static_cast<UINT32>(rect.bottom - rect.top));
    render_target_->Resize(size);
}

void RendererD2D::Render(HWND window, float dpi_scale, const ModuleInfo& module_info) {
    if (!EnsureRenderTarget(window)) {
        return;
    }

    RECT client{};
    GetClientRect(window, &client);
    const float width = static_cast<float>(client.right - client.left);
    const float height = static_cast<float>(client.bottom - client.top);

    auto* target = render_target_.Get();
    target->BeginDraw();
    target->Clear(Rgb(247, 249, 252));

    const float margin = Scale(24.0f, dpi_scale);
    const float top = Scale(24.0f, dpi_scale);
    const float header_height = Scale(96.0f, dpi_scale);
    const float status_height = Scale(28.0f, dpi_scale);

    D2D1_RECT_F header = D2D1::RectF(margin, top, width - margin, top + header_height);
    target->FillRoundedRectangle(
        D2D1::RoundedRect(header, Scale(8.0f, dpi_scale), Scale(8.0f, dpi_scale)), accent_brush_.Get());

    DrawTextBlock(target,
                  module_info.title,
                  D2D1::RectF(header.left + Scale(24.0f, dpi_scale),
                              header.top + Scale(18.0f, dpi_scale),
                              header.right - Scale(24.0f, dpi_scale),
                              header.top + Scale(50.0f, dpi_scale)),
                  text_brush_.Get(),
                  title_format_.Get());

    DrawTextBlock(target,
                  module_info.subtitle,
                  D2D1::RectF(header.left + Scale(24.0f, dpi_scale),
                              header.top + Scale(56.0f, dpi_scale),
                              header.right - Scale(24.0f, dpi_scale),
                              header.bottom - Scale(12.0f, dpi_scale)),
                  muted_brush_.Get(),
                  body_format_.Get());

    D2D1_RECT_F panel = D2D1::RectF(margin,
                                   header.bottom + Scale(18.0f, dpi_scale),
                                   width - margin,
                                   std::max(header.bottom + Scale(180.0f, dpi_scale), height - status_height - margin));
    target->FillRoundedRectangle(
        D2D1::RoundedRect(panel, Scale(8.0f, dpi_scale), Scale(8.0f, dpi_scale)),
        text_brush_.Get());
    target->DrawRoundedRectangle(
        D2D1::RoundedRect(panel, Scale(8.0f, dpi_scale), Scale(8.0f, dpi_scale)),
        border_brush_.Get(),
        Scale(1.0f, dpi_scale));

    DrawTextBlock(target,
                  L"Module status",
                  D2D1::RectF(panel.left + Scale(24.0f, dpi_scale),
                              panel.top + Scale(20.0f, dpi_scale),
                              panel.right - Scale(24.0f, dpi_scale),
                              panel.top + Scale(52.0f, dpi_scale)),
                  accent_brush_.Get(),
                  title_format_.Get());

    DrawTextBlock(target,
                  BuildBulletText(module_info),
                  D2D1::RectF(panel.left + Scale(24.0f, dpi_scale),
                              panel.top + Scale(64.0f, dpi_scale),
                              panel.right - Scale(24.0f, dpi_scale),
                              panel.bottom - Scale(24.0f, dpi_scale)),
                  muted_brush_.Get(),
                  body_format_.Get());

    D2D1_RECT_F status = D2D1::RectF(0, height - status_height, width, height);
    target->FillRectangle(status, border_brush_.Get());
    DrawTextBlock(target,
                  module_info.status,
                  D2D1::RectF(Scale(14.0f, dpi_scale),
                              height - status_height + Scale(6.0f, dpi_scale),
                              width - Scale(14.0f, dpi_scale),
                              height),
                  muted_brush_.Get(),
                  status_format_.Get());

    const HRESULT hr = target->EndDraw();
    if (hr == D2DERR_RECREATE_TARGET) {
        DiscardDeviceResources();
    }
}

std::wstring RendererD2D::BuildBulletText(const ModuleInfo& module_info) const {
    std::wstring text;
    for (const std::wstring& bullet : module_info.bullets) {
        if (!text.empty()) {
            text += L"\n";
        }
        text += L"- ";
        text += bullet;
    }
    return text;
}

bool RendererD2D::EnsureRenderTarget(HWND window) {
    if (render_target_) {
        return true;
    }

    RECT rect{};
    GetClientRect(window, &rect);
    D2D1_SIZE_U size =
        D2D1::SizeU(static_cast<UINT32>(rect.right - rect.left), static_cast<UINT32>(rect.bottom - rect.top));

    HRESULT hr = d2d_factory_->CreateHwndRenderTarget(D2D1::RenderTargetProperties(),
                                                      D2D1::HwndRenderTargetProperties(window, size),
                                                      render_target_.GetAddressOf());
    if (FAILED(hr)) {
        return false;
    }

    render_target_->CreateSolidColorBrush(Rgb(18, 26, 38), text_brush_.GetAddressOf());
    render_target_->CreateSolidColorBrush(Rgb(92, 106, 124), muted_brush_.GetAddressOf());
    render_target_->CreateSolidColorBrush(Rgb(31, 118, 210), accent_brush_.GetAddressOf());
    render_target_->CreateSolidColorBrush(Rgb(222, 228, 236), border_brush_.GetAddressOf());
    return text_brush_ && muted_brush_ && accent_brush_ && border_brush_;
}

void RendererD2D::DrawTextBlock(ID2D1HwndRenderTarget* target,
                                const std::wstring& text,
                                const D2D1_RECT_F& rect,
                                ID2D1Brush* brush,
                                IDWriteTextFormat* format) {
    target->DrawTextW(text.c_str(), static_cast<UINT32>(text.size()), format, rect, brush);
}

}  // namespace mytools
