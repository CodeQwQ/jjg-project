import defaultSettings from '@/settings'

const title = defaultSettings.title || '公路工程交（竣）工验收检测系统'

export default function getPageTitle(pageTitle) {
  if (pageTitle) {
    return `${pageTitle} - ${title}`
  }
  return `${title}`
}
