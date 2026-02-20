from datetime import datetime
from typing import Optional
from loguru import logger

try:
    from pydantic import BaseModel, Field
except ImportError:
    # Fallback for environments where pydantic binary components are missing (e.g. Python 3.14)
    logger.warning("Pydantic not fully available, using basic class fallback.")
    class Field:
        def __init__(self, default=None, default_factory=None, **kwargs):
            self.default = default
            self.default_factory = default_factory
    
    class BaseModel:
        def __init__(self, **data):
            # 1. Apply values from data
            for k, v in data.items():
                setattr(self, k, v)
            
            # 2. Look at class-level attributes to find Field defaults
            for k, v in self.__class__.__dict__.items():
                if isinstance(v, Field) and k not in data:
                    if v.default_factory:
                        setattr(self, k, v.default_factory())
                    elif v.default is not Ellipsis:
                        setattr(self, k, v.default)
                    else:
                        setattr(self, k, None) # Or raise error for required fields if needed
        
        def model_dump(self):
            return self.__dict__

class Trade(BaseModel):
    """
    Represents a single stock trade record.
    """
    symbol: str = Field(..., description="The stock ticker symbol (e.g., 2344)")
    stock_name: Optional[str] = Field(None, description="The Chinese name of the stock")
    side: str = Field(..., description="BUY or SELL")
    quantity: float = Field(..., gt=0, description="Number of shares traded")
    price: float = Field(..., gt=0, description="Price per share")
    date: datetime = Field(default_factory=datetime.now, description="Transaction timestamp")
    total_amount: Optional[float] = Field(None, description="Total transaction value")
    broker: Optional[str] = Field("Unknown", description="Broker name")
    order_id: Optional[str] = Field(None, description="Unique reference number")

    def __init__(self, **data):
        super().__init__(**data)
        if self.total_amount is None:
            self.total_amount = round(self.quantity * self.price, 2)
