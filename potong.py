input_file = "QR_SETTLE_360004_000829_250822_ACQ"
output_file = "QR_SETTLE_360004_000829_250822_ACQ_hsl"

with open(input_file, "r", encoding="utf-8") as f:
    lines = f.readlines()

with open(output_file, "w", encoding="utf-8") as f:
    f.writelines(lines[:221450])
