# Реализация

## Обработчик создания заказа

Обработчик проверяет состав заказа и сохраняет его одной транзакцией (листинг [@lst:create]). Если хотя бы одной позиции нет на складе, заказ не создается.

Листинг: Обработчик POST /orders {#lst:create}

```python
@app.post("/orders")
def create_order(payload: OrderIn) -> OrderOut:
    with db.transaction():
        order = repository.add_order(payload)
    return OrderOut.from_model(order)
```

Полный текст модуля вынесен в приложение [@app:code].
