local start_at = 7
local counters = {0, 0, 0} -- Счетчики для уровней 1, 2, 3
local first_h1_found = false

function Header(el)
  local level = el.level
  
  -- Игнорируем заголовки выше 3 уровня
  if level > 3 then return el end

  if level == 1 then
    if not first_h1_found then
      -- Это самый первый заголовок 1 уровня в теле документа
      first_h1_found = true
      counters[1] = start_at
    else
      counters[1] = counters[1] + 1
    end
    -- Сбрасываем подуровни при новом разделе 1 уровня
    counters[2] = 0
    counters[3] = 0
    
  elseif level == 2 then
    counters[2] = counters[2] + 1
    counters[3] = 0
    
  elseif level == 3 then
    counters[3] = counters[3] + 1
  end

  -- Формируем строку номера: 10, 10.1, 10.1.1
  local num_str = tostring(counters[1])
  if level >= 2 then num_str = num_str .. "." .. tostring(counters[2]) end
  if level >= 3 then num_str = num_str .. "." .. tostring(counters[3]) end

  -- Принудительно задаем атрибут number, который Pandoc использует для DOCX
  el.attributes["number"] = num_str
  
  return el
end