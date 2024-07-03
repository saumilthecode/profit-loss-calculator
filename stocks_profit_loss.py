import os
import pandas as pd

def calculate_trade_profit_loss(buy_price, sell_price, quantity, date, stock_name):
    buy_commission = 0.55
    sell_commission = 0.59
    total_buy_cost = (buy_price * quantity) + buy_commission
    total_sell_value = (sell_price * quantity) - sell_commission
    profit_loss = total_sell_value - total_buy_cost
    return date, stock_name, profit_loss

def main():
    # Path to your Excel file
    file_path = 'trades.xlsx'


    # Check if file exists and is accessible
    if not os.path.isfile(file_path):
        print(f"File not found: {file_path}")
        return

    try:
        # Read trades from Excel file
        trades_df = pd.read_excel(file_path)
        print("Excel file read successfully.")
        print("Columns in the DataFrame:", trades_df.columns)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    daily_profit_loss = {}

    for index, trade in trades_df.iterrows():
        try:
            date, stock_name, profit_loss = calculate_trade_profit_loss(
                trade['Buy Price'], trade['Sell Price'], trade['Quantity'], trade['Date'], trade['Stock Name']
            )
        except KeyError as e:
            print(f"Column not found: {e}")
            print("Available columns:", trades_df.columns)
            return

        if date not in daily_profit_loss:
            daily_profit_loss[date] = {}
        if stock_name not in daily_profit_loss[date]:
            daily_profit_loss[date][stock_name] = 0
        daily_profit_loss[date][stock_name] += profit_loss

    # Print daily profit/loss
    for date, stocks in daily_profit_loss.items():
        for stock_name, profit_loss in stocks.items():
            print(f"Date: {date}, Stock: {stock_name}, Profit/Loss: {profit_loss:.2f}")

    # Save daily profit/loss to Excel file
    result = []
    for date, stocks in daily_profit_loss.items():
        for stock_name, profit_loss in stocks.items():
            result.append({'Date': date, 'Stock Name': stock_name, 'Profit/Loss': profit_loss})

    daily_profit_loss_df = pd.DataFrame(result)
    daily_profit_loss_df.to_excel('daily_profit_loss.xlsx', index=False)

if __name__ == "__main__":
    main()
