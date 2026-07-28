# 简历生成数据格式

将最终确认的简历内容保存为 UTF-8 JSON，再交给
`scripts/build_resume_docx.py`。JSON 只保存当前任务数据，不要放入技能目录。

## 顶层字段

- `name`：姓名。
- `contact`：包含 `phone`、`email`、`location`。
- `sections`：按最终展示顺序排列的栏目。
- `photo_path`：可选，本地照片绝对路径。脚本会裁剪替换模板照片槽。

联系方式、经历日期等单值字段可写成字符串；需要黄标时写成：

```json
{"text": "【请补充：手机号】", "pending": true}
```

## 栏目类型

### `bullets`

用于自我评价等没有经历抬头的条目：

```json
{
  "title": "自我评价",
  "kind": "bullets",
  "items": [
    {
      "title": "模型研发能力",
      "body": "使用 PyTorch 完成图像分类模型的训练、评测与误差分析闭环。"
    }
  ]
}
```

`title` 只写 4–10 个字，不要自行添加 `【★】`。

### `entries`

用于教育、实习、工作、课外和项目经历：

```json
{
  "title": "项目经历",
  "kind": "entries",
  "items": [
    {
      "left": "图像分类模型",
      "middle": "PyTorch｜ResNet18",
      "right": {"text": "【请补充：项目时间】", "pending": true},
      "bullets": [
        {
          "title": "多数据集评测",
          "body": "完成三个数据集的训练与评测，验证准确率达到 91.45%。"
        }
      ]
    }
  ]
}
```

教育背景可将 `bullets` 设为空数组。

### `body`

用于个人技能等标签加正文的内容：

```json
{
  "title": "个人技能",
  "kind": "body",
  "items": [
    {
      "label": "算法与框架：",
      "body": "Python、PyTorch、CNN、模型训练与评测。"
    }
  ]
}
```

## 正文中的黄色待补项

正文可从字符串改为片段数组。只把完整待补提示设为黄色：

```json
[
  {"text": "处理并校验 ", "highlight": false},
  {"text": "【请补充：数据规模】", "highlight": true},
  {"text": " 条记录，将周期缩短至 4 小时。", "highlight": false}
]
```

黄色片段必须完整使用 `【请补充：具体信息】`，不能只写“待补数据”。
