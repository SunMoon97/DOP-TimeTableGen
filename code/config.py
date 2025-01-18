import os
import random
import json
import pandas as pd

class TimetableConfig:
    ORDERED_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    DAYS_MWF = ["Monday", "Wednesday", "Friday"]
    DAYS_TT = ["Tuesday", "Thursday"]
    TIME_SLOTS = [f"{hour}:00" for hour in range(8, 13)] + ["1:00"] + [f"{hour}:00" for hour in range(2, 6)]
    TIME_SLOTS_WITHOUT_8AM = [slot for slot in TIME_SLOTS if slot != "8:00"]
