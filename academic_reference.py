"""Small reviewed topic briefs, not a claim of live fact checking.

Only use a brief for a matching user topic. URLs record provenance; the author
must not invent a wider bibliography or treat generated draft text as evidence.
Reviewed 2026-10-20.
"""
import re

COMPUTER_MEMORY_SOURCES = (
    "https://learn.microsoft.com/en-us/windows/win32/memory/working-set",
    "https://learn.microsoft.com/en-us/windows-server/virtualization/hyper-v/dynamic-memory",
    "https://www.kingston.com/en/ssd/desktop-laptop-workstations",
    "https://learn.microsoft.com/en-us/troubleshoot/windows-server/performance/troubleshoot-performance-problems-in-windows",
)

COMPUTER_MEMORY_BRIEF = """
Проверенные уточнения по памяти компьютера и накопителям:
• Различай критерии сравнения. SSD обычно выигрывает у HDD по задержке доступа и скорости.
Выбор HDD для большого архива может быть связан с ценой за объём, а не с большей скоростью.
Не пиши сокращение вроде «SSD быстрее для мелких файлов, HDD — для больших»: оно меняет смысл.
Скорость конкретной операции зависит от моделей, интерфейса, нагрузки и характера доступа.
• Page fault означает, что обращение нельзя обслужить текущим отображением страницы.
Hard page fault требует чтения из файла подкачки или отображённого файла. Soft page fault
разрешается без чтения с накопителя, например когда страница уже находится в физической памяти.
Не утверждай, что любой page fault означает недостаток RAM или чтение с диска.
• Память виртуальной машины может выделяться статически или динамически, в зависимости
от гипервизора и настроек. Например, Hyper-V поддерживает Dynamic Memory.
Не утверждай, что каждая ВМ всегда занимает фиксированный объём или что всегда нужно
физической RAM больше суммы максимальных лимитов всех ВМ. Учитывай текущую потребность,
режим выделения, нагрузку хоста и запас для одновременной работы.
• Паузы интерфейса сами по себе не доказывают нехватку RAM или неисправность накопителя.
Сопоставляй загрузку CPU, занятую и доступную память, активность подкачки и диска,
нагрузку процессов во время задержки. Предлагай проверку причины, а не покупку по одному симптому.
Используй эти уточнения только в соответствующих разделах, не вставляй справку целиком
и не повторяй её в каждом разделе. Сохрани условия и исключения при сокращении текста.
Если пользователь дал противоречащее утверждение, явно объясни расхождение;
не выдавай ошибку за факт и не меняй дословную цитату незаметно.
Источники для этих уточнений (не добавляй непроверенные названия, даты или другие ссылки):
"""


def reference_guidance(data):
    """Use substantive inputs only; metadata and genre alone must not select a brief."""
    text = "\n".join(str(data.get(k) or "") for k in ("topic", "user_text", "extra"))
    if not re.search(r"\b(?:RAM|SSD|HDD|DRAM|NVMe)\b|оперативн\w*\s+памят|"
                     r"накопител\w*\s+компьютер|виртуальн\w*\s+памят|файл\w*\s+подкачк",
                     text, re.I):
        return ""
    return COMPUTER_MEMORY_BRIEF + "\n".join(COMPUTER_MEMORY_SOURCES)
