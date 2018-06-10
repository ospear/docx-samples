# サンプル2: htmltoword
#
# Usage: bundle exec ruby app/sample2/sample2.rb
#
# 結果:
#   テキスト: OK
#   表: OK
#   画像: NG, 未対応らしい

require 'htmltoword'

html = File.read('app/sample2/template.html')
# document = Htmltoword::Document.create(html)
Htmltoword::Document.create_and_save(html, 'app/sample2/result.docx')
