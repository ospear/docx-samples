# サンプル1: gem docx
#
# Usage: bundle exec ruby app/sample1/sample1.rb
#
# ブックマークを頼りにテキスト挿入は簡単に可能
# 画像や表の動的生成はdocx gem自体には機能がない
# NokogiriのAPIを上手く扱えば各ノード操作はできるが。。。
#
# 元となるテンプレートで書式がほぼ決定している場合には使えそう
# ブックマークをあらかじめセットしてテンプレートとなるファイルを用意
# 画像はword/media配下のファイル差し替えで対応可能?
# 表は列行数が決まってればあとはブックマークでテキスト挿入は可能

require 'docx'

# 開く
doc = Docx::Document.open('app/sample1/template.docx')

# ブックマークを利用してテキスト挿入
doc.bookmarks['bookmark_1'].insert_text_after("Hello world.")

# ブックマークを利用して複数行テキスト挿入
doc.bookmarks['bookmark_2'].insert_multiple_lines(['Hello', 'World', 'foo'])

# パラグラフ削除
doc.paragraphs.each do |p|
  p.remove! if p.to_s =~ /TODO/
end

# 変数展開(Nokogori)
vars = {
  'var1' => 'hoge',
  'var2' => 'fuga',
  'var3' => 'piyo',
  'var4' => 'foo'
}
doc.doc.xpath('//w:t').each do |el|
  vars.each do |key, val|
    el.content = el.content.gsub(/####{key}###/, val)
    el.content = el.content.gsub(/@@#{key}@@/, val)
    el.content = el.content.gsub(/\$#{key}/, val)
    el.content = el.content.gsub(/\${#{key}}/, val)
  end
end

# 画像差し替え
# TODO: 対応

# 保存
doc.save('app/sample1/sample1.docx')
