class BrokerageMonthlyDeals:  # (Base)

    def __init__(self, sheet, start_pos: int, stop_pos: int):
        self.class_name = self.__class__.__name__
        self.sheet = sheet
        self.start_row: int = start_pos
        self.stop_row: int = stop_pos
