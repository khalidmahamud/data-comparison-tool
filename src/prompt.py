import re


def inject_variables(content: str, variables: dict[str, str]):
  def replacer(match):
    key = match.group(1).strip()
    return str(variables.get(key, f"{{{{{key}}}}}"))

  return re.sub(r"{{\s*(\w+)\s*}}", replacer, content)


def read_file(path: str):
  with open(path, "r", encoding="utf-8") as file:
    content = file.read()

    if path.endswith('.md'):
      content = re.sub(r"<!--.*?-->", "", content, flags=re.DOTALL)

    return content.strip()


generate_prompt = read_file("./prompts/regenerate_hadis_prompt.md")
translate_arabic_to_bangla_prompt = read_file("./prompts/translate_arabic_to_bangla.md")
