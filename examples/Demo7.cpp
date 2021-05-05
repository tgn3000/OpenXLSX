#include <OpenXLSX.hpp>
#include <iostream>
#include <string>

using std::cout;
using std::endl;

struct Order {
    std::string m_orderDateTime;

    std::array<std::string, 2> m_coinPair;

    enum class BuySell : uint8_t {
        Buy = 0,
        Sell = 1
    };

    enum class TransactionFeeType : uint8_t {
        USDT = 0,
        BNB = 1
    };

    enum class TransactionStatus : int {
        UnknownStatus = -1,
        Canceled = 0,
        Filled = 1
    };

    BuySell m_buySell = BuySell::Buy;

    double m_orderPrice;
    double m_orderAmount;
    double m_avgTradingPrice;
    double m_quantityFilled;

    TransactionStatus m_orderStatus = TransactionStatus::Filled;

    // trading
    std::string m_tradeDateTime;
    double m_tradingPrice;
    double m_totalPrice;
    double m_transactionFee;
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
    << "order status = " << static_cast<int>(order.m_quantityFilled) << ", "
    << "trade time = " << order.m_tradeDateTime << ", "
    << "trade price = " << order.m_tradingPrice << ", "
    << "transaction fee = " << order.m_transactionFee << ", "
    << "transaction fee type = " << static_cast<int>(order.m_transactionFeeType) << ", "
    << std::endl;
    return os;
}

std::pair<int, Order::TransactionStatus> getRowStatus(OpenXLSX::XLWorksheet& wks, int rowIndex) {
    const char columnChar = 'A' + (9 - 1); // 'I'
    const std::string cellString = columnChar + std::to_string(rowIndex);
    const OpenXLSX::XLCell& cell = wks.cell(cellString);
    switch (cell.valueType()) {
        case OpenXLSX::XLValueType::String: {
            auto statusString = cell.value().get<std::string>();
            if(statusString == "Filled") {
                std::cout << "Cell " << cellString << " is filled\n";
                return {rowIndex+2, Order::TransactionStatus::Filled};
            }
            break;
        }
        default: {}
    }
    return {rowIndex, Order::TransactionStatus::UnknownStatus};
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
        if(orderStatus == Order::TransactionStatus::Filled) {
            Order order;
            order.m_orderDateTime = wks.cell('A' + rowString).value().get<std::string>();
            order.m_coinPair = {wks.cell('B' + rowString).value().get<std::string>(), ""};
            order.m_buySell = wks.cell('C' + rowString).value().get<std::string>() == "BUY" ? Order::BuySell::Buy : Order::BuySell::Sell;
            order.m_orderPrice = std::stod(wks.cell('D' + rowString).value().get<std::string>());

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