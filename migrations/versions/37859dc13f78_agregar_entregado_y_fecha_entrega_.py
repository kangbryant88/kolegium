"""agregar entregado y fecha_entrega_coordinacion a proyecto_aprendizaje

Revision ID: 37859dc13f78
Revises: 29dbbc785d74
Create Date: 2026-08-26 16:12:43.884099

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '37859dc13f78'
down_revision = '29dbbc785d74'
branch_labels = None
depends_on = None


def upgrade():
    # Nota: el autogenerate de Alembic detectó también otros cambios de
    # esquema (drift previo no relacionado, ej. grado.usuario_id,
    # usuario.cv_path) que NO pertenecen a esta migración -se dejaron fuera
    # a propósito para no borrar columnas/datos ajenos a esta tarea-.
    with op.batch_alter_table('proyecto_aprendizaje', schema=None) as batch_op:
        batch_op.add_column(sa.Column('entregado', sa.Boolean(), nullable=False, server_default=sa.text('0')))
        batch_op.add_column(sa.Column('fecha_entrega_coordinacion', sa.DateTime(), nullable=True))


def downgrade():
    with op.batch_alter_table('proyecto_aprendizaje', schema=None) as batch_op:
        batch_op.drop_column('fecha_entrega_coordinacion')
        batch_op.drop_column('entregado')
