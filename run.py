import io
import os
import sys
import json
from pprint import pprint as pp

import click
import httpx
import xlsxwriter

http = httpx.Client()

def make_file_getter(filename):
    folder = "cached_downloads/"
    os.makedirs(folder, exist_ok=True)
    def func(read_cache=True, write_cache=True, print_stuff=True):
        data = None
        cached = False

        if read_cache:
            try:
                with open(folder + filename) as f:
                    print(" from cache...", end="", flush=True)
                    data = json.load(f)
            except (FileNotFoundError, json.decoder.JSONDecodeError):
                pass
            else:
                cached = True

        if not data:
            print(" downloading...", end="", flush=True)
            response = http.get(f"https://mtgjson.com/api/v5/{filename}")
            data = json.loads(response.content)

        if not cached and write_cache:
            with open(folder + filename, "w") as f:
                json.dump(data, f)

        return data
    return func


def get_cards(set_name):
    return make_file_getter(f"{set_name}.json")()["data"]["cards"]

get_prices = make_file_getter("AllPricesToday.json")
get_sets = make_file_getter("SetList.json")

def write_cards_to_spreadsheet(filename, cards, prices, set_codes, limit=None, print_stuff=True):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    worksheet.set_column("A:A", 25)
    worksheet.set_column("B:B", 30)
    worksheet.set_column("C:C", 10)
    worksheet.set_column("D:D", 10)
    worksheet.set_column("E:E", 10)
    worksheet.set_column("F:F", 15)
    worksheet.set_column("G:G", 10)
    worksheet.set_column("H:H", 30)
    worksheet.set_column("I:I", 25)

    worksheet.write(0, 0, "Image")
    worksheet.write(0, 1, "Name")
    worksheet.write(0, 2, "Rarity")
    worksheet.write(0, 3, "Colors")
    worksheet.write(0, 4, "Price [USD]")
    worksheet.write(0, 5, "Foil Price [USD]")
    worksheet.write(0, 6, "Playability")
    worksheet.write(0, 7, "Set")
    worksheet.write(0, 8, "Block")

    # print("Name, Rarity, Colors, Price [USD], Foil Price [USD]")
    max_i = min(len(cards), limit or float("inf"))
    digits = len(str(max_i))
    progress = None

    min_rank = min(c["edhrecRank"] for c in cards if "edhrecRank" in c)
    max_rank = max(c["edhrecRank"] for c in cards if "edhrecRank" in c)

    for i, card in enumerate(cards):
        del card["foreignData"]

        if limit and i >= limit:
            break

        if progress:
            sys.stdout.write("\b" * len(progress))

        if print_stuff:
            progress = f" {i+1:>{digits}}/{max_i} ({(i+1)/max_i*100:>5.1f}%)"
            print(progress, end="", flush=True)

        try:
            price = prices["data"][card["uuid"]]
            try:
                retail_price = price["paper"]["tcgplayer"]["retail"]
            except KeyError:
                retail_price = price["paper"]["cardkingdom"]["retail"]

            try:
                normal_price = list(retail_price["normal"].values())[0]
            except KeyError:
                normal_price = ""
            try:
                foil_price = list(retail_price["foil"].values())[0]
            except KeyError:
                foil_price = ""
            # print(f'{card["name"]}, {card["rarity"]}, {" ".join(card["colors"])}, {normal_price}, {foil_price}')

            row = i+1
            worksheet.set_row(row, 200)
            worksheet.write(row, 1, card["name"])
            worksheet.write(row, 2, card["rarity"])
            worksheet.write(row, 3, " ".join(card["colors"]))
            worksheet.write(row, 4, normal_price)
            worksheet.write(row, 5, foil_price)

            if "edhrecRank" in card:
                playability = 5 - round(card["edhrecRank"] / max_rank * 5)
                worksheet.write(row, 6, f"{playability}/5")

            worksheet.write(row, 7, set_codes[card["setCode"]]["name"])
            worksheet.write(row, 8, set_codes[card["setCode"]]["block"])

            scryfall_id = card["identifiers"]["scryfallId"]
            image_url = f"https://cards.scryfall.io/small/front/{scryfall_id[0]}/{scryfall_id[1]}/{scryfall_id}.jpg"
            response = http.get(image_url)
            worksheet.embed_image(row, 0, image_url, {"image_data": io.BytesIO(response.content)})



        except Exception as e:
            print("Failed on card:")
            pp(card)
            print("Price:")
            print(price)
            print("Image:")
            print(response)
            raise e

    workbook.close()

@click.command()
@click.argument("name")
@click.option("--print_sets", is_flag=True, help="Print The Available sets.")
@click.option("--card_limit", type=int, default=None, help="Limit to this many cards.")
def main(name, print_sets, card_limit):
    if print_sets:
        print("Printing list of sets.")
    else:
        print(f"Making spreadsheet for {name}")
    file_name = f"{name}.xlsx"

    print("Loading sets...", end="", flush=True)
    sets = get_sets()
    print(" done.")

    set_codes = {}
    for s in sets["data"]:
        if s["type"] != "expansion":
            continue
        s.pop("decks", None)
        s.pop("languages", None)
        s.pop("sealedProduct", None)
        s.pop("translations", None)
        if print_sets:
            print(f"{s['name']:<40} {s.get('block', ''):<40} {s['releaseDate']}")
        if s.get("name", "").lower() == name.lower() or s.get("block", "").lower() == name.lower():
            set_codes[s["code"]] = s

    if print_sets:
        return

    if not set_codes:
        print(f"No sets found for '{name}'")
        return
    else:
        print(f"Found {len(set_codes)} sets for {name}: {', '.join(s['name'] for s in set_codes.values())}")

    cards = []

    for code, s in set_codes.items():
        print(f"Loading cards for {s['name']}...", end="", flush=True)
        cards.extend(get_cards(code))
        print(" done.")

    print("Loading prices...", end="", flush=True)
    prices = get_prices()
    print(" done.")

    print("Processing cards...", end="", flush=True)
    write_cards_to_spreadsheet(file_name, cards, prices, set_codes, limit=card_limit)
    print(" done.")
    print(f"Wrote {file_name}")


if __name__ == "__main__":
    main()