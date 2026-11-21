-- Only the two requested substitutions. The DOCX reader/writer handles layout.
function Str(el)
  if el.text == "{{title_en}}" then
    return pandoc.Inlines({pandoc.Str("Forms"), pandoc.Space(), pandoc.Str("of"), pandoc.Space(), pandoc.Str("Evidence")})
  end
  if el.text == "{{author}}" then return pandoc.Str("Anonymous") end
end
