#include <OpenXLSX.hpp>
#include <iostream>
#include <string>

using std::cout;
using std::endl;

struct Order {
    enum class BuySell : uint8_t {
        Buy = 0,
        Sell = 1
    };

    enum class OrderStatus : int {
        UnknownStatus = -1,
        Canceled = 0,
        Filled = 1
    };

    enum class TransactionFeeType : int {
        UnknownCoin = -1,
        USDT = 0,
        BNB = 1
    };

    // order row
    std::string m_orderDateTime;
    std::array<std::string, 2> m_coinPair;
    BuySell m_buySell = BuySell::Buy;

    double m_orderPrice;
    double m_orderAmount;
    double m_avgTradingPrice;
    OrderStatus m_orderStatus = OrderStatus::Filled;

    // trading row
    std::string m_tradeDateTime;
    double m_tradingPrice;
    double m_quantityFilled;
    double m_totalPrice;
    std::string m_transactionFee;
    TransactionFeeType m_transactionFeeType = TransactionFeeType::USDT;
};

inline std::ostream& operator<<(std::ostream& os, const Order& order) {
    os 
    << "order time = " << order.m_orderDateTime << ", " 
    << "rate pair: (" << order.m_coinPair[0] << ", " << order.m_coinPair[1] << "), "
    << "order price = " << order.m_orderPrice << ", "
    << "amount = " << order.m_orderAmount << ", "
    << "avg trading price = " << order.m_avgTradingPrice << ", "
    << "quantity filled = " << order.m_quantityFilled << ", "
    << "order status = " << static_cast<int>(order.m_orderStatus) << ", "
    << "trade time = " << order.m_tradeDateTime << ", "
    << "trade price = " << order.m_tradingPrice << ", "
    << "transaction fee = " << order.m_transactionFee << ", "
    << "transaction fee type = " << static_cast<int>(order.m_transactionFeeType) << ", "
    << std::endl;
    return os;
}

std::pair<int, Order::OrderStatus> getRowStatus(OpenXLSX::XLWorksheet& wks, int rowIndex) {
    const char columnChar = 'A' + (9 - 1); // 'I'
    const std::string cellString = columnChar + std::to_string(rowIndex);
    const OpenXLSX::XLCell& cell = wks.cell(cellString);
    switch (cell.valueType()) {
        case OpenXLSX::XLValueType::String: {
            auto statusString = cell.value().get<std::string>();
            if(statusString == "Filled") {
                std::cout << "Cell " << cellString << " is filled\n";
                return {rowIndex+2, Order::OrderStatus::Filled};
            }
            break;
        }
        default: {}
    }
    return {rowIndex, Order::OrderStatus::UnknownStatus};
}

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #07: Gary's driver program\n";
    cout << "********************************************************************************\n";

    OpenXLSX::XLDocument doc;
    doc.open("./transaction.xlsx");
    auto wks = doc.workbook().worksheet("sheet1");

    std::string columnString = "A";
    std::string rowString = "1";
    std::string cellString = columnString + rowString;

    // auto PrintCell = [](const OpenXLSX::XLCell& cell) {
    //     cout << "Cell type is ";

    //     switch (cell.valueType()) {
    //         case OpenXLSX::XLValueType::Empty:
    //             cout << "XLValueType::Empty";
    //             break;

    //         case OpenXLSX::XLValueType::Float:
    //             cout << "XLValueType::Float and the value is " << cell.value().get<double>() << endl;
    //             break;

    //         case OpenXLSX::XLValueType::Integer:
    //             cout << "XLValueType::Integer and the value is " << cell.value().get<int64_t>() << endl;
    //             break;

    //         case OpenXLSX::XLValueType::String:
    //             cout << "XLValueType::String and the value is " << cell.value().get<std::string>() << endl;
    //             break;

    //         default:
    //             cout << "Unknown";
    //     }
    //     cout << endl;
    // };

    constexpr int numberColumns = 'I' - 'A' + 1;
    int rowIndex = 1, columnIndex = 1;
    while(rowIndex < 128) {
        rowString = std::to_string(rowIndex);

        // for(int columnIndex = 1; columnIndex <= numberColumns; ++columnIndex) {
            // const char columnChar = 'A' + (columnIndex - 1);
            // columnString = columnChar;
            // cellString = columnString + rowString;
            // cout << "Cell " << cellString << ": ";

        auto [rowIndexEnd, orderStatus] = getRowStatus(wks, rowIndex);
        if(orderStatus == Order::OrderStatus::Filled) {
            Order order;
            rowString = std::to_string(rowIndex);

            // order row
            order.m_orderDateTime   = wks.cell('A' + rowString).value().get<std::string>();
            order.m_coinPair        = {wks.cell('B' + rowString).value().get<std::string>(), ""};
            order.m_buySell         = wks.cell('C' + rowString).value().get<std::string>() == "BUY" ? Order::BuySell::Buy : Order::BuySell::Sell;
            order.m_orderPrice      = std::stod(wks.cell('D' + rowString).value().get<std::string>());
            order.m_orderAmount     = std::stod(wks.cell('E' + rowString).value().get<std::string>());
            order.m_avgTradingPrice = std::stod(wks.cell('F' + rowString).value().get<std::string>());
            order.m_orderStatus     = wks.cell('I' + rowString).value().get<std::string>() == "Filled" ? Order::OrderStatus::Filled : Order::OrderStatus::Canceled;

            // trading row
            rowIndex += 2;
            rowString = std::to_string(rowIndex);
            order.m_tradeDateTime   = wks.cell('B' + rowString).value().get<std::string>();
            order.m_tradingPrice    = std::stod(wks.cell('C' + rowString).value().get<std::string>());
            order.m_quantityFilled  = std::stod(wks.cell('D' + rowString).value().get<std::string>());
            order.m_totalPrice      = std::stod(wks.cell('E' + rowString).value().get<std::string>());
            order.m_transactionFee  = wks.cell('F' + rowString).value().get<std::string>();

            if(order.m_transactionFee.substr(order.m_transactionFee.size()-4) == "USDT") {
                order.m_transactionFeeType = Order::TransactionFeeType::USDT;
            }
            else if(order.m_transactionFee.substr(order.m_transactionFee.size()-3) == "BNB") {
                order.m_transactionFeeType = Order::TransactionFeeType::BNB;
            }
            else {}
            cout << order;
        }


            rowIndex = rowIndexEnd;

            // const OpenXLSX::XLCell& cell = wks.cell(cellString);
            // PrintCell(cell);
        // }
        ++rowIndex;
    }

    // doc.resetCalcChain();
    doc.saveAs("./transaction_out.xlsx");
    doc.close();
    return 0;
}