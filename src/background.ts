(() => {
    class TeamsBackend {
        readonly contextMenuId = "teams-see-more-context-menu-id";
        readonly messageName = "teams-see-more-message";
        matchedTab: { [key: number]: boolean; } = {};
        contextMenuCreated: boolean = false;
        activeTabId: number = -1;
        isUrlMatched(url?: string) {
            if (url != null && url.indexOf("https://teams.microsoft.com/") >= 0) {
                return true;
            }
            return false;
        }

        updateContextMenu(): void {
            if (this.activeTabId === -1) {
                return;
            }
            const tabId = this.activeTabId;
            const isMatched = this.matchedTab[tabId];
            // console.log(`updateContextMenus - ${isMatched}`);
            if (isMatched) {
                if (!this.contextMenuCreated) {
                    this.contextMenuCreated = true;
                    // console.log(`Tab ${tabId} is now teams`);
                    chrome.contextMenus.create({
                        id: this.contextMenuId,
                        title: "Teams See More", // chrome.i18n.getMessage("extName"),
                        contexts: ["page"]
                    }, () => {
                        if (chrome.runtime.lastError) {
                            // console.log(chrome.runtime.lastError.message);
                            return;
                        }
                    });
                }
            }
            else {
                if (this.contextMenuCreated) {
                    this.contextMenuCreated = false;
                    // console.log(`Tab ${tabId} is no longer teams`);
                    chrome.contextMenus.remove(this.contextMenuId, () => {
                        if (chrome.runtime.lastError) {
                            // console.log(chrome.runtime.lastError.message);
                            return;
                        }
                    });
                }
            }
        }

        constructor() {
            // Fired when a tab is updated.
            chrome.tabs.onUpdated.addListener((tabId: number, changeInfo: chrome.tabs.TabChangeInfo, tab: chrome.tabs.Tab) => {
                if (tab == null) {
                    return;
                }
                if (changeInfo.url != null) {
                    // console.log("onUpdated: " + tab.id, JSON.stringify(changeInfo));
                    this.matchedTab[tabId] = this.isUrlMatched(changeInfo.url);
                }
                if (changeInfo.status === "complete") {
                    // console.log("onUpdated: " + tab.id, JSON.stringify(changeInfo), tab.url);
                    this.matchedTab[tabId] = this.isUrlMatched(tab.url);
                    if (this.matchedTab[tabId]) {
                        chrome.scripting.executeScript({
                            target: { tabId },
                            files: ["foreground.js"]
                        }).then((value: chrome.scripting.InjectionResult[]) => {
                            if (chrome.runtime.lastError) {
                                // console.log(chrome.runtime.lastError.message);
                                return;
                            }
                            this.updateContextMenu();
                        }).catch((reason: any) => {
                            if (chrome.runtime.lastError) {
                                // console.log(chrome.runtime.lastError.message);
                                return;
                            }
                        })
                    }
                    else {
                        this.updateContextMenu();
                    }
                }
            });

            // Fires when the active tab in a window changes.
            // Note that the tab's URL may not be set at the time this event fired, but you can listen to onUpdated events to be notified when a URL is set.
            chrome.tabs.onActivated.addListener((activeInfo: chrome.tabs.TabActiveInfo) => {
                if (activeInfo != null && activeInfo.tabId != null) {
                    this.activeTabId = activeInfo.tabId;
                    this.updateContextMenu();
                }
            });

            // Fired when a tab is created.
            // Note that the tab's URL may not be set at the time this event fired, but you can listen to onUpdated events to be notified when a URL is set.
            chrome.tabs.onCreated.addListener((tab: chrome.tabs.Tab) => {
                if (tab != null && tab.id != null) {
                    this.matchedTab[tab.id] = false;
                }
            });

            // Fired when a tab is closed.
            chrome.tabs.onRemoved.addListener((tabId: number, removeInfo: chrome.tabs.TabRemoveInfo) => {
                delete this.matchedTab[tabId];
            });

            // Fired when the currently focused window changes.
            // Will be chrome.windows.WINDOW_ID_NONE if all chrome windows have lost focus.
            chrome.windows.onFocusChanged.addListener((windowId: number) => {
                chrome.tabs.query({ active: true, currentWindow: true }, (tabs: chrome.tabs.Tab[]) => {
                    if (tabs.length > 0) {
                        const tab = tabs[0];
                        if (tab.id != null) {
                            this.activeTabId = tab.id;
                            this.updateContextMenu();
                        }
                    }
                });
            });

            // Fired when a context menu item is clicked.
            chrome.contextMenus.onClicked.addListener((info: chrome.contextMenus.OnClickData, tab?: chrome.tabs.Tab) => {
                // console.log('context menu clicked');
                // console.log(info);
                // console.log(tab);
                if (tab == null || tab.id == null) {
                    return;
                }
                if (info.menuItemId === this.contextMenuId) {
                    chrome.tabs.sendMessage(
                        tab.id,
                        {
                            message: this.messageName
                        },
                        (response) => {
                            if (chrome.runtime.lastError) {
                                return;
                            }
                            if (response != null && response.message === "success") {
                                // console.log("button#see_more succeeded.");
                            }
                        }
                    );
                    return true;
                }
            });
        }
    };
    new TeamsBackend();
})();