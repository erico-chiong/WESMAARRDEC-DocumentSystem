from django import template

register = template.Library()

@register.filter
def user_roles(user):
    roles = []
    if user.admin:
        roles.append("Admin")
    if user.user:
        roles.append("User")
    if user.researcher:
        roles.append("Researcher")
    if user.secretariat:
        roles.append("Secretariat")
    if user.stakeholder:
        roles.append("Stakeholder")
    return ", ".join(roles)
