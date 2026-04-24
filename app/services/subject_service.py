from pathlib import Path
import re

def clean_file_name(file_path: str) -> str:
    file_name = Path(file_path).stem

    #troca os _ e - por espaço

    file_name = file_name.replace("_", " ").replace("-", " ")

    #tira os espaços duplicados
    file_name = re.sub(r"\s+", " ", file_name).strip()

    return file_name


def generate_subject(client_name: str, attachments: list[str]) -> str:
    client_name = (client_name or " ").strip()

    if not client_name:
        return "Envio de arquivos"
    
    if not attachments:
        return f"{client_name} - Envio de arquivos"
    
    first_file_name = clean_file_name(attachments[0])

    if len(attachments) == 1:
        return f"{client_name} - {first_file_name}"
    
    extra_count = len(attachments) - 1
    return  f"{client_name} - {first_file_name} e mais {extra_count} arquivo(s)" 