# Приложение: Исходный код модуля заказов {#app:code}

Листинг: Модуль orders.py

```python
"""Order endpoints."""

def total(items):
	return sum(item.price * item.quantity for item in items)
```
