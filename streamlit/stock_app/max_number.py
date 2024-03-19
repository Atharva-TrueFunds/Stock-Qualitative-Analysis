def max_number(nums1, nums2, k):
    def max_single_number(nums, k):
        stack = []
        drop = len(nums) - k
        for num in nums:
            while drop and stack and stack[-1] < num:
                stack.pop()
                drop -= 1
            stack.append(num)
        return stack[:k]

    def merge(nums1, nums2):
        return [max(nums1, nums2).pop(0) for _ in range(len(nums1) + len(nums2))]

    max_num = []
    for i in range(max(0, k - len(nums2)), min(k, len(nums1)) + 1):
        max_num = max(
            max_num, merge(max_single_number(nums1, i), max_single_number(nums2, k - i))
        )
    return max_num


nums1 = [3, 4, 6, 5]
nums2 = [-9, 1, 2, 5, 8, 3]
k = 5
print(max_number(nums1, nums2, k))

nums1 = [7]
nums2 = [6, 0, 4]
k = 4
print(max_number(nums1, nums2, k))

nums1 = [3, 9]
nums2 = [8, 9]
k = 3
print(max_number(nums1, nums2, k))
