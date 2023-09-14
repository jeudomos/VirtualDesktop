﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using WindowsDesktop.Interop.Proxy;
using WindowsDesktop.Utils;

namespace WindowsDesktop.Interop.Build22621;

public class VirtualDesktopNotificationService : ComWrapperBase<IVirtualDesktopNotificationService>, IVirtualDesktopNotificationService
{
    private readonly ComWrapperFactory _factory;

    internal VirtualDesktopNotificationService(ComInterfaceAssembly assembly, ComWrapperFactory factory)
        : base(assembly, CLSID.VirtualDesktopNotificationService)
    {
        this._factory = factory;
    }
 
    public IDisposable Register(IVirtualDesktopNotification notification)
    {
        /* Could not find the IVirtualDesktopNotification for this build, do nothing */
        return Disposable.Create(() => { });
    }

    public abstract class EventListenerBase
    {
        internal ComWrapperFactory Factory { get; set; } = null!;

        internal IVirtualDesktopNotification Notification { get; set; } = null!;

        protected void CreatedCore(object pDesktop)
            => this.Notification.VirtualDesktopCreated(this.Wrap(pDesktop));

        protected void DestroyBeginCore(object pDesktopDestroyed, object pDesktopFallback)
            => this.Notification.VirtualDesktopDestroyBegin(this.Wrap(pDesktopDestroyed), this.Wrap(pDesktopFallback));

        protected void DestroyFailedCore(object pDesktopDestroyed, object pDesktopFallback)
            => this.Notification.VirtualDesktopDestroyFailed(this.Wrap(pDesktopDestroyed), this.Wrap(pDesktopFallback));

        protected void DestroyedCore(object pDesktopDestroyed, object pDesktopFallback)
            => this.Notification.VirtualDesktopDestroyed(this.Wrap(pDesktopDestroyed), this.Wrap(pDesktopFallback));

        protected void MovedCore(object p0, object pDesktop, int nIndexFrom, int nIndexTo)
            => this.Notification.VirtualDesktopMoved(this.Wrap(pDesktop), nIndexFrom, nIndexTo);
        
        protected void RenamedCore(object pDesktop, HString chName)
            => this.Notification.VirtualDesktopRenamed(this.Wrap(pDesktop), chName);

        protected void ViewChangedCore(object view)
            => this.Notification.ViewVirtualDesktopChanged(this.Factory.ApplicationView(view).Interface);

        protected void CurrentChangedCore(object pDesktopOld, object pDesktopNew)
            => this.Notification.CurrentVirtualDesktopChanged(this.Wrap(pDesktopOld), this.Wrap(pDesktopNew));

        protected void WallpaperChangedCore(object pDesktop, HString chPath)
            => this.Notification.VirtualDesktopWallpaperChanged(this.Wrap(pDesktop), chPath);

        private IVirtualDesktop Wrap(object desktop)
            => this.Factory.VirtualDesktop(desktop).Interface;
    }
}
