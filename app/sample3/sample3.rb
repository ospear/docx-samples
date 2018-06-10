# サンプル3: pandoc
#
# Prepare: brew install pandoc
# Usage: bundle exec ruby app/sample3/sample3.rb > app/sample3/sample3.docx
#
# なにか書式不正があるのか警告が出る。無理やり開いてもあまり綺麗ではない。。。
# 画像もパスを同設定すればいいのか不明

require 'pandoc-ruby'

html = File.read('app/sample3/template.html')
@converter = PandocRuby.new(html, from: :html, to: :docx)
puts @converter.convert
