function Header(el)
  if el.level == 1 then
    local sectionbreak = '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>'
    return {
      pandoc.RawBlock('openxml', sectionbreak),
      el
    }
  end
end