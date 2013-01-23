ReVIEW::Compiler.defblock :subtitle, 0
ReVIEW::Compiler.defblock :author, 0

class ReVIEW::HTMLBuilder
  def subtitle(lines)
    puts '<div class="subtitle">' + lines.join("\n") + '</div>'
  end

  def author(lines)
    puts '<div class="author">' +
      lines.map {|x| "<p class=\"author\">#{x}</p>" }.join('') + '</div>'
  end
end
