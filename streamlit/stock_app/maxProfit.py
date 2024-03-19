# On each day, you may decide to buy and/or sell the stock.
# You can only hold at most one share of the stock at any time.
# However, you can buy it then immediately sell it on the same day.
# Find and return the maximum profit you can achieve.


class Solution(object):
    def maxProfit(self, prices):
        """
        :type prices: List[int]
        :rtype: int
        """
        if not prices:
            return 0

        max_profit = 0
        for i in range(1, len(prices)):
            if prices[i] < 0:
                pass
            elif prices[i - 1] < 0:
                pass
            elif prices[i] > prices[i - 1]:
                max_profit += prices[i] - prices[i - 1]
        return max_profit


solution = Solution()
prices = [-5, 7, 1, 5, 4, 5, 6, 4]
print(solution.maxProfit(prices))
