/**
  MIT License

  Copyright (c) 2021 Danny Graham

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.
**/

const { basename, extname, dirname } = require('path')
const { readFileSync, writeFileSync } = require('fs')
const { TemplateHandler } = require('easy-template-x')

const uuid = () => Math.random().toString(36).substr(2, 9)

module.exports = function (RED) {
  function docx(config) {
    RED.nodes.createNode(this, config)
    const self = this
    this.on("input", async (msg) => {
      const handler = new TemplateHandler()
      const templateFile = readFileSync(`/data/${config.templatePath}`)
      const templateFileName = basename(`/data/${config.templatePath}`, extname(config.templatePath))
      const doc = await handler.process(templateFile, msg.payload.docx)

      msg.payload.docx = { ...msg.payload.docx, templateFileName }
      if (config.outputPath.indexOf('${') > -1) {
        for (let key of Object.keys(msg.payload)) {
          config.outputPath = config.outputPath.replace('${' + key + '}', msg.payload[key])
        }
        for (let key of Object.keys(msg.payload.docx)) {
          config.outputPath = config.outputPath.replace('${' + key + '}', msg.payload.docx[key])
        }
      }

      if (config.outputPath.indexOf('docx') === -1) {
        config.outputPath = `${config.outputPath}/${templateFileName}_${config.id}_${uuid()}.docx`
      }

      writeFileSync(`/data/${config.outputPath}`, doc)
      msg.payload.docx = { ...msg.payload.docx, ...config, templateFileName }
      self.send(msg)
    })
  }

  RED.nodes.registerType("docx", docx)
}
