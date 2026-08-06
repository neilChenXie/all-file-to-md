import path from "path"
import { fileURLToPath } from "url"

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const skillsDir = path.resolve(__dirname, "../../skills")

/**
 * V1 旧版格式插件：直接导出函数 (input, options?) => Promise<Hooks>
 *
 * V1 加载器先尝试 PluginModule 格式 { id, server() }，若未匹配则回退到旧版格式。
 * 此插件仅注册 skills 路径，无需 id 或 server 包装，旧版格式更简洁。
 *
 * config 钩子将 skills 目录注入运行时配置对象，经 V1→V2 迁移后由 ConfigSkillPlugin 发现。
 */
export default async (input, options) => {
  return {
    config: async (config) => {
      config.skills = config.skills || {}
      config.skills.paths = config.skills.paths || []
      if (!config.skills.paths.includes(skillsDir)) {
        config.skills.paths.push(skillsDir)
      }
    },
  }
}