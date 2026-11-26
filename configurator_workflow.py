"""Block identity and validation for the configurator workflow."""
from dataclasses import dataclass
from uuid import uuid4

from methodology_v2 import number

BLOCKS = "1. Блоки"
PRICES = "2. Цены"
SMR = "3. СМР"
INDIRECT = "4. Инфраструктура и косвенные"
RESULTS = "5. Итоги"
VOLUMES = "Объёмы"
RATES = "Коробка A–E"
PARKING = "Паркинг"
GZ = "G и Z"
BLOCK_ID = "ID урбан-блока"


@dataclass(frozen=True)
class Issue:
    section: str
    text: str
    block: str = ""
    subsection: str = ""
    group: str = ""

    @property
    def label(self):
        return f"{self.block} · {self.text}" if self.block else self.text


def block_identity(state, names):
    previous = state.get("block_ids", {})
    result = {}
    used = set()
    for name in names:
        identity = previous.get(name)
        if not isinstance(identity, str) or not identity or identity in used:
            identity = str(uuid4())
        result[name] = identity
        used.add(identity)
    state["block_ids"] = result
    return result


def reconcile_blocks(state, previous, edited, names, stores):
    """Apply simultaneous name edits to the existing calculation stores."""
    by_id = {identity: name for name, identity in previous.items()}
    renamed = {}
    identities = {}
    for identity, name in zip(edited, names):
        if isinstance(identity, str) and identity in by_id and identity not in identities.values():
            renamed[by_id[identity]] = name
        else:
            identity = str(uuid4())
        identities[name] = identity
    if previous != identities:
        old_names = set(previous)
        for key in stores:
            source = state.get(key, {})
            state[key] = {
                renamed.get(name, name): value for name, value in source.items()
                if name not in old_names or name in renamed
            }
        state["block_ids"] = identities
    return identities


def pair_issue(result, block, quantity, rate, label, section, subsection="", group="",
               quantity_section=BLOCKS, quantity_subsection="", *, require_quantity=True):
    qty = number(quantity)
    value = number(rate)
    if qty > 0 and value <= 0:
        result.append(Issue(section, f"{label} — заполните цену или ставку", block, subsection, group))
    elif require_quantity and value > 0 and qty <= 0:
        result.append(Issue(quantity_section, f"{label} — заполните объём", block, quantity_subsection))
