const fs = require('node:fs')
const path = require('node:path')
const crypto = require('node:crypto')
const yaml = require('js-yaml')
const pkg = require('../package.json')

async function verify() {
  const dir = path.resolve(__dirname, '..', 'release', pkg.version)
  const info = yaml.load(fs.readFileSync(path.join(dir, 'latest.yml'), 'utf8'))
  const expected = `mulk-yonetim-setup-${pkg.version}.exe`
  if (info.version !== pkg.version || info.path !== expected || info.files?.length !== 1 || info.files[0].url !== expected) {
    throw new Error('latest.yml sürümü veya kurulum dosyası adı paketle uyuşmuyor.')
  }
  const file = path.join(dir, expected)
  const hash = crypto.createHash('sha512')
  for await (const chunk of fs.createReadStream(file)) hash.update(chunk)
  const digest = hash.digest('base64')
  if (digest !== info.sha512 || digest !== info.files[0].sha512 || fs.statSync(file).size !== info.files[0].size) {
    throw new Error('Kurulum dosyasının boyutu veya SHA512 özeti latest.yml ile uyuşmuyor.')
  }
  fs.accessSync(`${file}.blockmap`)
  const feed = yaml.load(fs.readFileSync(path.join(dir, 'win-unpacked/resources/app-update.yml'), 'utf8'))
  if (feed.provider !== 'github' || feed.owner !== 'ggokhan58' || feed.repo !== 'mulkyonetim') {
    throw new Error('Paketin güncelleme adresi yanlış.')
  }
  console.log(`Doğrulandı: ${expected}, SHA512, latest.yml, blockmap ve GitHub adresi.`)
}
verify().catch(error => { console.error(error); process.exitCode = 1 })
