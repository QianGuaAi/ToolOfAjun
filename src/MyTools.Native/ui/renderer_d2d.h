#pragma once

#include <string>

#include <d2d1.h>
#include <dwrite.h>
#include <wrl/client.h>

#include "modules/module_registry.h"

namespace mytools {

class RendererD2D {
public:
    bool Initialize();
    void DiscardDeviceResources();
    void Resize(HWND window);
    void Render(HWND window, float dpi_scale, const ModuleInfo& module_info);

private:
    bool EnsureRenderTarget(HWND window);
    void DrawTextBlock(ID2D1HwndRenderTarget* target,
                       const std::wstring& text,
                       const D2D1_RECT_F& rect,
                       ID2D1Brush* brush,
                       IDWriteTextFormat* format);
    std::wstring BuildBulletText(const ModuleInfo& module_info) const;

    Microsoft::WRL::ComPtr<ID2D1Factory> d2d_factory_;
    Microsoft::WRL::ComPtr<IDWriteFactory> dwrite_factory_;
    Microsoft::WRL::ComPtr<ID2D1HwndRenderTarget> render_target_;
    Microsoft::WRL::ComPtr<ID2D1SolidColorBrush> text_brush_;
    Microsoft::WRL::ComPtr<ID2D1SolidColorBrush> muted_brush_;
    Microsoft::WRL::ComPtr<ID2D1SolidColorBrush> accent_brush_;
    Microsoft::WRL::ComPtr<ID2D1SolidColorBrush> border_brush_;
    Microsoft::WRL::ComPtr<IDWriteTextFormat> title_format_;
    Microsoft::WRL::ComPtr<IDWriteTextFormat> body_format_;
    Microsoft::WRL::ComPtr<IDWriteTextFormat> status_format_;
};

}  // namespace mytools
