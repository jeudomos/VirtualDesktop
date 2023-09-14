using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows.Media.Animation;
using WindowsDesktop.Interop.Proxy;

namespace WindowsDesktop.Interop.Build10240;

internal class VirtualDesktopManagerInternal : ComWrapperBase<IVirtualDesktopManagerInternal>, IVirtualDesktopManagerInternal
{
    private readonly ComWrapperFactory _factory;

    public VirtualDesktopManagerInternal(ComInterfaceAssembly assembly, ComWrapperFactory factory)
        : base(assembly, CLSID.VirtualDesktopManagerInternal)
    {
        this._factory = factory;
    }

    public IEnumerable<IVirtualDesktop> GetDesktops()
    {
        var array = this.InvokeMethod<IObjectArray>();
        if (array == null) yield break;

        var count = array.GetCount();
        var vdType = this.ComInterfaceAssembly.GetType(nameof(IVirtualDesktop));

        for (var i = 0u; i < count; i++)
        {
            var ppvObject = array.GetAt(i, vdType.GUID);
            yield return new VirtualDesktop(this.ComInterfaceAssembly, ppvObject);
        }
    }

    public IVirtualDesktop GetCurrentDesktop()
        => this.InvokeMethodAndWrap();

    public IVirtualDesktop GetAdjacentDesktop(IVirtualDesktop pDesktopReference, AdjacentDesktop uDirection)
    {
        // => this.InvokeMethodAndWrap(Args(((VirtualDesktop)pDesktopReference).ComObject, uDirection));
        var desktops=GetDesktops().ToArray();
        var current = GetCurrentDesktop();
        var currentId = current.GetID();
        var i = 0;
        while(i < desktops.Length && currentId != desktops[i].GetID()) { i++; }
        if (i < desktops.Length) { 
            if (i == 0) return desktops[desktops.Length - 1];
            return desktops[i - 1];
        }
        return current;
    }

    public IVirtualDesktop FindDesktop(Guid desktopId)
        => this.InvokeMethodAndWrap(Args(desktopId));

    public IVirtualDesktop CreateDesktop()
        => this.InvokeMethodAndWrap();

    public void SwitchDesktop(IVirtualDesktop desktop)
        => this.InvokeMethod(Args(((VirtualDesktop)desktop).ComObject));

    public void RemoveDesktop(IVirtualDesktop pRemove, IVirtualDesktop pFallbackDesktop)
        => this.InvokeMethod(Args(((VirtualDesktop)pRemove).ComObject, ((VirtualDesktop)pFallbackDesktop).ComObject));

    public void MoveViewToDesktop(IntPtr hWnd, IVirtualDesktop desktop)
        => this.InvokeMethod(Args(this._factory.ApplicationViewFromHwnd(hWnd).ComObject, ((VirtualDesktop)desktop).ComObject));

    public void SetDesktopName(IVirtualDesktop desktop, string name)
        => throw new NotSupportedException();

    public void SetDesktopWallpaper(IVirtualDesktop desktop, string path)
        => throw new NotSupportedException();

    public void UpdateWallpaperPathForAllDesktops(string path)
        => throw new NotSupportedException();

    private VirtualDesktop InvokeMethodAndWrap(object?[]? parameters = null, [CallerMemberName] string methodName = "")
        => new(this.ComInterfaceAssembly, this.InvokeMethod<object>(parameters, methodName) ?? throw new Exception("Failed to get IVirtualDesktop instance."));
}
