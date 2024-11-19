# A variable where we can save the mapping from AI to AI name as parsed from the (SA2) Agenda
sa2_ai_names_mapping: dict[str, str] = {}


def ai_to_wi_str(ai_str: str) -> str:
    try:
        return sa2_ai_names_mapping[ai_str]
    except KeyError:
        return ''
    except Exception as e:
        print(f'Could not map AI {ai_str}: {e}')
        return ''
