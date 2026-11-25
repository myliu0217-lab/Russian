from pathlib import Path

path = Path('俄语1review.md')
text = path.read_text(encoding='utf-8')
start = text.index('### （5）食物与吃饭')
end = text.index('</table>', start) + len('</table>')
section = text[start:end]
text = text[:start] + text[end:]
path.write_text(text.rstrip() + '\n\n' + section + '\n', encoding='utf-8')
