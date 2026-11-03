// 点工具栏图标：打开视频管理页
chrome.action.onClicked.addListener(() => {
  chrome.tabs.create({ url: 'https://channels.weixin.qq.com/platform/post/list' });
});
