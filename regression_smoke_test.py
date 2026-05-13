"""
DyeFlow RS Regression Smoke Test
Run: python regression_smoke_test.py
Purpose: catch regressions before packaging a new version.
"""
import copy
import sys

try:
    from main import calc
except Exception as exc:
    print("FAIL: main.py import failed:", exc)
    sys.exit(1)

BASE_PROJECT = {
    "project_name": "Regression Demo",
    "company_name": "Demo Company",
    "process_type": "Dye_PES",
    "fabric_kg": 100,
    "flote": 10,
    "carry_over": 2,
    "fabric_status": "Dry",
    "cost_currency": "EUR",
    "machine": {
        "machine_name": "REG-100", "capacity_kg": 100, "drain_time_min": 5,
        "circulation_pump_power": 10, "pump_ratio": 1,
        "number_of_reel": 1, "reel_power": 1, "fan_power": 0, "fan_ratio": 0
    },
    "utilities": {
        "heating_source": "Natural Gas", "heating_capacity": 8250, "transfer_heat_loss": 0,
        "natural_gas_unit_price": 1, "water_unit_price": 1, "waste_water_unit_price": 1,
        "electric_unit_price": 1, "hourly_wage": 60, "number_of_workers": 1, "number_of_machine": 1
    },
    "steps": [{
        "filling_time": 5, "beginning_temp": 50, "heating_slope": 2, "final_temp": 90,
        "dwelling_time": 10, "cooling_gradient": 2, "cooling_temp": 70, "overflow_time": 0,
        "amount_of_flote": 1000, "flote_ratio": 10, "bath_count": 1, "drain": False,
        "chemicals": [{
            "supplier": "REG", "chemical": "Test Chemical", "company": "", "process": "",
            "begin_c": 50, "final_c": 50, "dose_min": 1, "dose_time": 1,
            "circulation_time": 0, "amount": 1, "unit": "g/l", "price": 10
        }]
    }]
}


def check(condition, message):
    if not condition:
        print("FAIL:", message)
        sys.exit(1)
    print("OK:", message)


def main():
    no_overflow = calc(copy.deepcopy(BASE_PROJECT))["dashboard"]

    project_overflow = copy.deepcopy(BASE_PROJECT)
    project_overflow["steps"][0]["overflow_time"] = 10
    with_overflow = calc(project_overflow)["dashboard"]

    # Overflow extra water = (10 / 5) * 1000 = 2000 L
    check(with_overflow["Overflow Extra Water L / batch"] == 2000, "Overflow extra water formula")
    check(with_overflow["Total Water L / batch"] == 3000, "Total water includes overflow")
    check(with_overflow["Base Process Water L / batch"] == 1000, "Base process water unchanged")

    # Overflow must not affect heating or chemical cost.
    check(with_overflow["Heating Cost / batch"] == no_overflow["Heating Cost / batch"], "Overflow excluded from heating")
    check(with_overflow["Chemical Cost / batch"] == no_overflow["Chemical Cost / batch"], "Overflow excluded from chemical")

    # Overflow time must affect electricity and labour.
    check(with_overflow["Electricity (kWh/batch)"] > no_overflow["Electricity (kWh/batch)"], "Overflow time included in electricity")
    check(with_overflow["Labour Cost / batch"] > no_overflow["Labour Cost / batch"], "Overflow time included in labour")

    # Filling and drain time must not affect electricity active time.
    base = copy.deepcopy(BASE_PROJECT)
    base["steps"][0]["drain"] = True
    base_short = calc(copy.deepcopy(base))["dashboard"]
    base["steps"][0]["filling_time"] = 50
    base["machine"]["drain_time_min"] = 50
    base_long = calc(copy.deepcopy(base))["dashboard"]
    check(base_short["Electricity (kWh/batch)"] == base_long["Electricity (kWh/batch)"], "Electricity excludes filling and drain time")


    # Dose_min interval after filling must be counted as active electricity time.
    dose_test = copy.deepcopy(BASE_PROJECT)
    dose_test["steps"][0]["dose_min"] = 0
    dose0 = calc(copy.deepcopy(dose_test))["dashboard"]
    dose_test["steps"][0]["chemicals"][0]["dose_min"] = 5
    dose5 = calc(copy.deepcopy(dose_test))["dashboard"]
    check(dose5["Active Electricity Time (min)"] > dose0["Active Electricity Time (min)"], "Dose_min interval included in electricity")

    # Wastewater must follow actually taken fresh water, not nominal bath water after carry-over.
    carry = copy.deepcopy(BASE_PROJECT)
    carry["steps"][0]["drain"] = True
    carry["steps"].append(copy.deepcopy(BASE_PROJECT["steps"][0]))
    carry["steps"][1]["drain"] = True
    carry["steps"][1]["overflow_time"] = 0
    carry_result = calc(carry)["dashboard"]
    # Dry fabric: first step 1000 L, second step (10-2)*100 = 800 L, total wastewater 1800 L.
    check(carry_result["Total Water L / batch"] == 1800, "Carry-over reduces fresh water after drained first bath")
    check(carry_result["Wastewater L / batch"] == 1800, "Wastewater limited to actually taken water")

    # Drain=False creates no wastewater; Drain=True discharges the bath and overflow.
    check(with_overflow["Wastewater L / batch"] == 0, "No wastewater when drain is false")
    drained = copy.deepcopy(project_overflow)
    drained["steps"][0]["drain"] = True
    drained_result = calc(drained)["dashboard"]
    check(drained_result["Wastewater L / batch"] == 3000, "Wastewater counted only on drained step")

    # If previous step is not drained, next step takes no fresh water, but later drain discharges held bath.
    two_step = copy.deepcopy(BASE_PROJECT)
    two_step["steps"].append(copy.deepcopy(BASE_PROJECT["steps"][0]))
    two_step["steps"][0]["drain"] = False
    two_step["steps"][1]["drain"] = True
    two_step["steps"][1]["overflow_time"] = 0
    two_step_result = calc(two_step)["dashboard"]
    check(two_step_result["Total Water L / batch"] == 1000, "No new water if previous step was not drained")
    check(two_step_result["Wastewater L / batch"] == 1000, "Held bath discharged on later drained step")

    print("PASS: Regression smoke test completed successfully.")


if __name__ == "__main__":
    main()
