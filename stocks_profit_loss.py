import pandas as pd

def calculate_trade_profit_loss(buy_price, sell_price, quantity):
    buy_commission = 0.55
    sell_commission = 0.59
    total_buy_cost = (buy_price * quantity) + buy_commission
    total_sell_value = (sell_price * quantity) - sell_commission
    profit_loss = total_sell_value - total_buy_cost
    return profit_loss

def main():
    # Path to your Excel file
    file_path = 'trades.xlsx'

    try:
        # Read trades from Excel file
        trades_df = pd.read_excel(file_path)
        print("Excel file read successfully.")

        # Clean column names by stripping leading/trailing spaces
        trades_df.columns = trades_df.columns.str.strip()
        print("Columns in the DataFrame after stripping:", trades_df.columns)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    trades_df['Profit/Loss'] = trades_df.apply(lambda trade: calculate_trade_profit_loss(trade['Buy Price'], trade['Sell Price'], trade['Quantity']), axis=1)

    # Calculate accumulated profit/loss for each day
    trades_df['Accumulated Profit/Loss'] = trades_df.groupby('Date')['Profit/Loss'].cumsum()

    # Save the updated DataFrame back to Excel
    output_file = 'updated_trades.xlsx'
    trades_df.to_excel(output_file, index=False)
    print(f"Updated trades data saved to {output_file}")

if __name__ == "__main__":
    main()
