//? 分析比写代码更重要
//? 以空串为分割，将数组中的元素进行分组拼接，返回由拼接后的字串组成的新数组
//? [a,'',b,c,'',d] => [a,bc,d]
//? 使用 reduce 操作数组，返回一个数组

//? 以函数作为最小构建块，匿名函数是最小执行单元
//? 思路：pre 一开始是空数组，作为 reduce 的结果值
//? 遍历到 Nullish 时，查看下一个元素是否有值，有值则向 pre 添加空串，便于后续聚合
//? 只有当前元素有值时才聚合
//?
const src = ['', '', 'a', undefined, 'b', 'z', 'c', '', 'd']
const result = src.reduce((pre, val, idx, arr) => {
  pre = [].concat(pre)
  let last = pre.length - 1

  // 判断 val 是不是空串
  if (val) {
    // 不是空串，且 pre 中有空串承接
    pre[last] = (pre[last] ?? '') + val
  } else if (arr[idx + 1]) {
    // 当前元素是空串，检视下一个元素是否为
    pre.length++
  }
  return pre
})

console.log(result)
