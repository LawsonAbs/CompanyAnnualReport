

def get_key_words():
    key_words_text="国内生产总值、反腐败、八项规定、高质量发展、供给侧、和谐社会、美丽中国、国际经济、非公有制经济、国内外经济形势、宏观调控、国民经济、货币政策、通货膨胀、国际经济环境、结构性改革、中国制造“2025”、一带一路、京津冀协同发展、海上丝绸之路、产能过剩、亚太自贸区、长江经济带、混合所有制经济、经济合作区、利率市场化"
    key_words = []
    key_words = key_words_text.split("、")    
    return key_words

if __name__ == "__main__":
    a = get_key_words()
    print(a)