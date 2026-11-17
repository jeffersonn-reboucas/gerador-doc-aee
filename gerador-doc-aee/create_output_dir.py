import os
from pathlib import Path

output_dir = Path("./rascunhos/teste_saida")
output_dir.mkdir(parents=True, exist_ok=True)
print(f"Output directory ensured: {output_dir}")